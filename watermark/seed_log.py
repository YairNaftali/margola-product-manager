"""One-time seed: mark every media ID currently live on the catalog as
already-watermarked, since the full 2026-08-12 production run already
watermarked all of them. Run once, never again (run_watermark_batch.py
appends to the same log on every future run)."""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from app import shopify_graphql

from run_watermark_batch import gql_with_retry, PRODUCTS_QUERY, LOG_PATH

ids = []
cursor = None
product_count = 0
while True:
    data = gql_with_retry(PRODUCTS_QUERY, {"after": cursor, "query": None})
    edges = data["products"]["edges"]
    for e in edges:
        product_count += 1
        if e["node"]["media"]["pageInfo"]["hasNextPage"]:
            print(f"WARNING: {e['node']['title']} has >100 images, not fully captured")
        for m in e["node"]["media"]["edges"]:
            ids.append(m["node"]["id"])
    if not data["products"]["pageInfo"]["hasNextPage"] or not edges:
        break
    cursor = edges[-1]["cursor"]

open(LOG_PATH, "w", encoding="utf-8").write(json.dumps(sorted(set(ids)), indent=2))
print(f"Seeded {len(set(ids))} media IDs across {product_count} products into {LOG_PATH}")
