# Product image watermarking

Watermarks live Shopify product photos with the Margola logo, in place, via
the Admin GraphQL API. Built 2026-08-11/12 (first full-catalog run: 620/620
products, 1,079 image attachments, 0 failures), rebuilt as reusable tooling
2026-08-13 so future new products/photos don't need this re-derived from
scratch.

## Locked settings

Position `center`, logo width 35% of image width, opacity 40%. This was
Yair's explicit call over an earlier per-category/corner-placement version —
center sits on top of the product and is harder to crop out.

**Known open issue, not fixed:** the navy-gray logo has weak contrast against
light/silver/crystal product photos (e.g. Rhinestone Squaredelles, crystal
Chaton Rose) — legible but noticeably fainter than on colored beads. If this
needs fixing later: bump opacity specifically for light backgrounds, or add a
subtle dark outline/shadow behind the logo.

## How to run it

```
cd ~/Downloads/Margola_Product_Manager_v4/watermark
source ../.venv/bin/activate

python3 run_watermark_batch.py --dry-run          # see what it *would* touch, changes nothing
python3 run_watermark_batch.py                    # process every product with unwatermarked images
python3 run_watermark_batch.py --collection <handle>   # scope to one collection
python3 run_watermark_batch.py --product-ids 123,456   # scope to specific products
python3 run_watermark_batch.py --limit 5           # stop after 5 products actually changed (spot-test)
```

Always run `--dry-run` first after onboarding a new category or big photo
batch, confirm the count looks right, then run for real.

## How it decides what's "new"

`watermarked_media_ids.json` in this folder is the source of truth — a flat
list of every Shopify media ID this tool has ever produced. On each run, it
fetches every product's current media and treats **any media ID not in that
file** as needing a watermark. This is what makes re-runs safe:

- A brand new product → all its images are unknown IDs → all get watermarked.
- A new photo added to an already-watermarked product → only that one new
  image is an unknown ID → only it gets watermarked, the rest of the
  product's media (and its position) is left untouched.
- Running the script twice in a row → second run finds nothing new, does
  nothing (confirmed via `--dry-run` against the live catalog after seeding).

The log was seeded once (`seed_log.py`, already run 2026-08-13) with the
1,079 media IDs live at the end of the original full run — do **not**
re-run `seed_log.py` except to recover from a lost/corrupted log file, since
it blindly marks whatever's currently live as "already watermarked," which
would be wrong if some live images genuinely aren't watermarked yet.

**If a photo gets replaced** (old one deleted, new unwatermarked one
uploaded under a fresh media ID) it will automatically get picked up and
watermarked on the next run — nothing extra needed.

## Mechanism (what actually happens per product)

Uses `productCreateMedia` + `productDeleteMedia` + `productReorderMedia` —
narrow, single-product mutations (deprecated in the schema but functional;
no non-deprecated narrow replacement exists for this). Deliberately **not**
`productSet`, which is a full-product upsert and too risky for an in-place
image swap (anything not explicitly included in the call risks being reset).

Per product with N new images: download each unique source image once
(dedup'd by URL — some products share the exact same underlying MediaImage
record, e.g. the "Chaton Rose" line), watermark it locally with `watermark.py`,
stage-upload it (`stagedUploadsCreate` → POST to the returned target →
`resourceUrl`), batch-create all N as new media in one call, batch-delete the
N old ones in one call, then reorder the **entire current media list** (not
just the new items) back into original position in one call — this last part
is what makes partial/incremental runs safe (see next section).

Rate limiting: every mutation call retries with exponential backoff on
`THROTTLED` errors; a short pause every 25 *updated* products.

## Bugs from the original run, already fixed in this version

1. **Manifest pagination cap** — the original backup query used
   `media(first: 20)`, silently truncating products with 21+ images (several
   Chaton Rose products). This version uses `media(first: 100)` and
   explicitly warns + skips reordering on any product still hitting that cap
   (`hasNextPage: true`) rather than silently mis-ordering it.
2. **Position-preservation bug on partial batches** — the original script
   only reordered the newly-created images into positions `0..N-1`, which
   displaced already-correct images on the 2 products it partially touched
   (had to be fixed by hand after the fact). This version reconstructs the
   *entire* product's media order every time — old kept media stays where it
   was, new media slots into the position its predecessor occupied — verified
   with a local unit test (fake product, 4 images, 2 "new") before trusting
   it against live data.

## Scope/safety

Only ever touches `product.media` (fetched per-product), never the general
Files library or theme assets — structurally guarantees the homepage hero,
Shop by Product grid photos, banners, and the logo file itself can never be
touched, since none of those are attached to a product record.

## Files in this folder

- `watermark.py` — core PIL watermarking function.
- `run_watermark_batch.py` — the runner described above.
- `seed_log.py` — one-time log seeder (already run, don't re-run casually).
- `watermarked_media_ids.json` — the "already done" log, source of truth for
  what needs processing. Keep this file, don't delete it.
- `run_progress.jsonl` — append-only log of every product-level result from
  every run (status, count, product name), for debugging/audit.

## Dead end, for context

Embedding copyright/EXIF metadata (Copyright/Rights/UsageTerms/WebStatement
fields via exiftool) was tested and confirmed **not viable** for the live
site — Shopify strips all such metadata and re-encodes the file on upload
(tested twice, independently, including a real upload→download→compare
round trip). Still open, lower priority: whether to embed this into local/
archival copies for off-Shopify use (print catalog, Etsy, wholesale) where
Shopify's stripping wouldn't apply.
