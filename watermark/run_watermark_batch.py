"""
Reusable watermark runner for Margola product images.

Watermarks any product image that hasn't been watermarked yet (tracked by
Shopify media ID in watermarked_media_ids.json), swaps it into the live
product in place, and records the new media IDs so re-runs are idempotent.

Safe to run repeatedly / incrementally: only touches media IDs it has never
seen before, so it naturally picks up new products, newly-added photos on
existing products, or replaced photos -- without re-watermarking anything
already done or corrupting image order on products it only partially touches.

See README.md in this folder for the full writeup (settings, mechanism,
known gotchas). Requires the project's .venv (Pillow + requests) and the
same .env Shopify credentials app.py uses.

Usage:
    python3 watermark/run_watermark_batch.py                 # scan whole catalog, process anything new
    python3 watermark/run_watermark_batch.py --collection rhinestone-banding
    python3 watermark/run_watermark_batch.py --product-ids 123456789,987654321
    python3 watermark/run_watermark_batch.py --dry-run        # report what would change, touch nothing
    python3 watermark/run_watermark_batch.py --limit 5        # cap to first 5 products with new images (testing)
"""
import argparse
import json
import os
import sys
import tempfile
import time
import urllib.parse

import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from app import shopify_graphql  # noqa: E402  (reuses the project's existing auth/GraphQL helper)

from watermark import apply_watermark  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
LOG_PATH = os.path.join(HERE, "watermarked_media_ids.json")
PROGRESS_PATH = os.path.join(HERE, "run_progress.jsonl")

BATCH_PAUSE_EVERY = 25
BATCH_PAUSE_SECONDS = 4
MAX_RETRIES = 6


def load_watermarked_ids():
    if os.path.exists(LOG_PATH):
        return set(json.load(open(LOG_PATH, encoding="utf-8")))
    return set()


def save_watermarked_ids(ids):
    open(LOG_PATH, "w", encoding="utf-8").write(json.dumps(sorted(ids), indent=2))


def gql_with_retry(query, variables=None):
    delay = 2
    for attempt in range(MAX_RETRIES):
        try:
            return shopify_graphql(query, variables)
        except RuntimeError as e:
            if "THROTTLED" in str(e) and attempt < MAX_RETRIES - 1:
                time.sleep(delay)
                delay = min(delay * 2, 30)
                continue
            raise


PRODUCTS_QUERY = """
query GetProducts($after:String, $query:String){
  products(first:50, after:$after, query:$query){
    edges{ cursor node{
      id title
      media(first:100){
        pageInfo{ hasNextPage }
        edges{ node{
          id alt mediaContentType
          ... on MediaImage{ image{ url width height } }
        } }
      }
    } }
    pageInfo{ hasNextPage }
  }
}
"""


def fetch_products(collection_handle=None, product_ids=None):
    if product_ids:
        gids = [pid if str(pid).startswith("gid://") else f"gid://shopify/Product/{pid}" for pid in product_ids]
        q = "query($ids:[ID!]!){ nodes(ids:$ids){ ... on Product{ id title media(first:100){ pageInfo{hasNextPage} edges{ node{ id alt mediaContentType ... on MediaImage{ image{ url width height } } } } } } } }"
        data = gql_with_retry(q, {"ids": gids})
        return [n for n in data["nodes"] if n]

    query_filter = None
    if collection_handle:
        query_filter = f"collection_id:{collection_handle}"  # not used directly; see note below

    products, cursor = [], None
    search_query = None
    if collection_handle:
        # Resolve handle -> id, then filter by collection membership via query string.
        cdata = gql_with_retry("query($h:String!){ collectionByHandle(handle:$h){ id } }", {"h": collection_handle})
        coll = cdata.get("collectionByHandle")
        if not coll:
            raise RuntimeError(f"No collection found for handle '{collection_handle}'")
        coll_num = coll["id"].split("/")[-1]
        search_query = f"collection_id:{coll_num}"

    while True:
        data = gql_with_retry(PRODUCTS_QUERY, {"after": cursor, "query": search_query})
        edges = data["products"]["edges"]
        products.extend(e["node"] for e in edges)
        if not data["products"]["pageInfo"]["hasNextPage"] or not edges:
            break
        cursor = edges[-1]["cursor"]
    return products


def download(url, dest_path):
    r = requests.get(url, timeout=60)
    r.raise_for_status()
    open(dest_path, "wb").write(r.content)


def staged_upload_and_create(product_gid, jpeg_path, alt_text):
    filename = os.path.basename(jpeg_path)
    size = os.path.getsize(jpeg_path)
    staged = gql_with_retry(
        """mutation($input:[StagedUploadInput!]!){ stagedUploadsCreate(input:$input){
             stagedTargets{ url resourceUrl parameters{ name value } } userErrors{ field message } } }""",
        {"input": [{
            "filename": filename, "mimeType": "image/jpeg", "resource": "IMAGE",
            "fileSize": str(size), "httpMethod": "POST",
        }]},
    )["stagedUploadsCreate"]
    if staged["userErrors"]:
        raise RuntimeError(staged["userErrors"])
    target = staged["stagedTargets"][0]
    form = {p["name"]: p["value"] for p in target["parameters"]}
    with open(jpeg_path, "rb") as f:
        resp = requests.post(target["url"], data=form, files={"file": (filename, f, "image/jpeg")}, timeout=120)
    resp.raise_for_status()
    return target["resourceUrl"], alt_text


def swap_product(product, needs_processing_ids, cache, dry_run=False):
    """needs_processing_ids: media IDs on this product not yet in the watermark log."""
    product_gid = product["id"]
    all_media = [e["node"] for e in product["media"]["edges"]]
    if product["media"]["pageInfo"]["hasNextPage"]:
        return {"status": "skipped_over_100_images", "product": product["title"]}

    to_process = [m for m in all_media if m["id"] in needs_processing_ids]
    if not to_process:
        return {"status": "nothing_new"}

    if dry_run:
        return {"status": "would_process", "count": len(to_process), "product": product["title"]}

    # Build final ordered list: for each original media slot, either the kept
    # (already-watermarked) node or the freshly-created replacement.
    new_media_inputs = []
    old_media_ids_to_delete = []
    id_to_new_source = {}

    for m in to_process:
        url = m["image"]["url"]
        source_key = url.split("?")[0]
        if source_key not in cache:
            with tempfile.TemporaryDirectory() as td:
                src = os.path.join(td, "src.jpg")
                out = os.path.join(td, "wm.jpg")
                download(url, src)
                apply_watermark(src, out)
                resource_url, _ = staged_upload_and_create(product_gid, out, m.get("alt") or "")
                cache[source_key] = resource_url
        id_to_new_source[m["id"]] = cache[source_key]
        new_media_inputs.append({
            "originalSource": cache[source_key],
            "alt": m.get("alt") or "",
            "mediaContentType": "IMAGE",
        })
        old_media_ids_to_delete.append(m["id"])

    created = gql_with_retry(
        """mutation($productId:ID!, $media:[CreateMediaInput!]!){
             productCreateMedia(productId:$productId, media:$media){
               media{ id } mediaUserErrors{ field message } } }""",
        {"productId": product_gid, "media": new_media_inputs},
    )["productCreateMedia"]
    if created["mediaUserErrors"]:
        raise RuntimeError(created["mediaUserErrors"])
    new_ids = [m["id"] for m in created["media"]]

    gql_with_retry(
        """mutation($productId:ID!, $mediaIds:[ID!]!){
             productDeleteMedia(productId:$productId, mediaIds:$mediaIds){
               deletedMediaIds mediaUserErrors{ field message } } }""",
        {"productId": product_gid, "mediaIds": old_media_ids_to_delete},
    )

    # Reconstruct full original order, substituting new ids for processed slots.
    kept_ids = [m["id"] for m in all_media if m["id"] not in needs_processing_ids]
    # all_media is already in the product's real current order; walk it and
    # substitute in place so images that weren't touched don't move.
    old_to_new = dict(zip(old_media_ids_to_delete, new_ids))
    final_order = [old_to_new.get(m["id"], m["id"]) for m in all_media if m["id"] in old_to_new or m["id"] in kept_ids]

    moves = [{"id": mid, "newPosition": str(i)} for i, mid in enumerate(final_order)]
    gql_with_retry(
        """mutation($id:ID!, $moves:[MoveInput!]!){
             productReorderMedia(id:$id, moves:$moves){ mediaUserErrors{ field message } } }""",
        {"id": product_gid, "moves": moves},
    )

    return {"status": "processed", "count": len(new_ids), "product": product["title"], "new_media_ids": new_ids}


def run_incremental(collection=None, product_ids=None, limit=None, dry_run=False, on_progress=None):
    """Core incremental scan+watermark loop -- shared by the CLI (main(), below)
    and the Product Manager app's automatic background trigger (app.py's
    watermark_run_background(), fired on every "Sync Collections" click so
    nobody has to run this script by hand). Only ever touches media IDs never
    seen before (see module docstring), so scanning the whole catalog on every
    call is cheap once most products are already done.

    on_progress(scanned, total, touched, product_title), if given, is called
    after every product (whether or not it needed anything) so a caller can
    surface live status.
    """
    products = fetch_products(collection_handle=collection, product_ids=product_ids)
    watermarked_ids = load_watermarked_ids()

    cache = {}
    progress = open(PROGRESS_PATH, "a", encoding="utf-8")
    touched = 0
    limit_hits = 0  # counts "processed" OR "would_process", so --limit also stops a --dry-run early
    total = len(products)

    for i, product in enumerate(products, start=1):
        media_nodes = [e["node"] for e in product["media"]["edges"]]
        needs = {m["id"] for m in media_nodes if m["id"] not in watermarked_ids}
        if needs:
            result = swap_product(product, needs, cache, dry_run=dry_run)
            result["product_id"] = product["id"]
            progress.write(json.dumps(result) + "\n")
            progress.flush()
            print(result)

            if result["status"] == "processed":
                watermarked_ids.update(result["new_media_ids"])
                save_watermarked_ids(watermarked_ids)  # save after every product so a crash loses nothing
                touched += 1
            if result["status"] in ("processed", "would_process"):
                limit_hits += 1

        if on_progress:
            on_progress(i, total, touched, product.get("title", ""))
        if limit and limit_hits >= limit:
            break
        if touched and touched % BATCH_PAUSE_EVERY == 0:
            time.sleep(BATCH_PAUSE_SECONDS)

    progress.close()
    return {"scanned": total, "touched": touched}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--collection", help="collection handle to scope to, e.g. rhinestone-banding")
    ap.add_argument("--product-ids", help="comma-separated numeric or gid product IDs")
    ap.add_argument("--limit", type=int, help="stop after N products that need processing (testing)")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    product_ids = args.product_ids.split(",") if args.product_ids else None
    result = run_incremental(collection=args.collection, product_ids=product_ids, limit=args.limit, dry_run=args.dry_run)
    print(f"\nScanned {result['scanned']} products, {result['touched']} updated.")


if __name__ == "__main__":
    main()
