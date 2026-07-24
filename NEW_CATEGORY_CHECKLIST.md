# Adding a new product category

Steps to onboard a new spreadsheet/category (seed beads, bugle beads, fire polished,
3 Cut, etc.), based on how 2 Cut Beads 11/0 was added.

## 1. Import the spreadsheet
- Drop it in and try `Import Spreadsheet`. If the sheet uses a merged two-row header
  (a group label in row 1, the real field name in row 2 -- like the 2 Cut Beads
  template), `parse_xlsx` already auto-detects that. If a column doesn't map, check
  `get_first(row, [...])` candidate lists in `parse_xlsx` (app.py) -- add the exact
  header text as another candidate rather than renaming the spreadsheet.
- If the category isn't in `data/spreadsheet_types.json` yet, add a new entry
  (matches on `sheet_match`/`filename_match` -- substrings checked against the
  compacted sheet tab name / lowercased filename, same semantics as the old
  hardcoded `infer()` branches). Set `collection`, `subcategory`, `bead_shape`,
  `color_type` (leave `""` if the sheet has its own explicit Color Type column --
  it gets picked up automatically), `shopify_category`, and **`sync`** -- the real
  Shopify collection(s) this category's approved products should be added to when
  "Sync Collections" runs. Ask Yair for the target collection when he sends the
  filelist for this category (this is what replaced the old blanket-`collection`
  sync behavior that caused the 2026-07-24 incident) -- leave `sync: []` until he
  confirms it, which makes Sync Collections a safe no-op for this type in the
  meantime.
  - Once added, the new type shows up automatically in the import page's
    "Spreadsheet type" dropdown -- pick it explicitly instead of relying on
    auto-detect if you want to be sure which recipe applies.
- If the desired product Title doesn't fit the default
  `{prefix} {descriptor} {size} {color}` pattern, give the type a `title_suffix`
  field -- see the `2cut-beads` entry for the alternate `{prefix} {size} {color}
  {suffix}` construction, including the "drop a redundant trailing BEADS" and
  whitespace-typo cleanup logic next to it in `parse_xlsx` (app.py).
- One piece of matching logic that couldn't move into the JSON: Roller Beads'
  transparent/opaque `color_type` detection (scans both sheet name and color name
  together) stays as a small special case in `infer()`, gated on `type_id ==
  "roller-beads"`.

## 2. Bead shape taxonomy
- `data/shopify_taxonomy_map.json` -> `"bead_shape"` needs an entry for the new
  category: `{"label": "...", "base": "<one of BEAD_SHAPE_BASE_GIDS>"}`.
- Before guessing a base taxonomy value, check whether Shopify already has a live
  metaobject for this shape (another dev may have added it by hand):
  `python3 -c "import app; [print(n['displayName'], n['handle']) for n in app.shopify_list_metaobjects('shopify--bead-shape')]"`
  -- if it exists, match its `taxonomy_reference` GID exactly (this is how "2 Cut
  Beads" was matched to the same base as the site's existing live "3 Cut").

## 3. Color taxonomy (swatch metafield)
- `data/shopify_taxonomy_map.json` -> `"color"` maps each raw color name to one of
  Shopify's fixed 19 taxonomy colors. **New categories will have zero entries here**
  -- check before a real export:
  ```
  python3 -c "
  import json
  tm = json.load(open('data/shopify_taxonomy_map.json'))['color']
  products = json.load(open('data/products.json'))
  missing = [p['color_name'] for p in products if p['color_name'] not in tm]
  print(len(missing), 'unmapped colors'); [print(' -', m) for m in missing]
  "
  ```
  Without a mapping, `MetaobjectResolver.color_handle()` silently skips the swatch
  for that product (empty cell in the CSV, no error) -- these color->taxonomy calls
  need a human judgment (Neil has historically made these calls), so don't guess-fill
  them yourself; flag the list and get them confirmed.

## 4. Photo resolver (drive matching)
Only needed once photos exist on a drive. Pattern from 2 Cut Beads:
1. Generate a filelist for the folder (tab-separated `size / mtime / path` --
   `stat -f` does **not** interpret `\t` in its format string on macOS; use
   `printf` instead, e.g.:
   `find "<folder>" -type f -exec sh -c 'printf "%s\t%s\t%s\n" "$(stat -f %z "$1")" "$(stat -f %m "$1")" "$1"' _ {} \; > ~/Downloads/<name>_filelist.txt`
2. Add that path to `DRIVE_FILELISTS`.
3. Add a bucket to `load_drive_index()` (a dict + an `elif "<folder marker>" in fl:`
   line), matching a distinctive substring of the folder path.
4. Add a small matcher function (see `_find_2cut_photo` for the pattern: try a few
   filename suffix variants against `color_number`) and wire it into
   `resolve_photo_from_drive()` with an `elif product.get("bead_shape") == "...":`
   branch.
- If the new category's real filenames also follow `{color_number}[-_]{size}.jpg`
  (likely, for other Preciosa Ornela categories), `_find_2cut_photo` can probably be
  reused directly or with a renamed/generalized copy instead of writing new matching
  logic from scratch.

## 5. Descriptions
No code involved -- hand-write them following the established voice (see
`czech_glass_beads_descriptions.txt` and `2cut_11-0_descriptions.txt` for the
reference layout/tone), one `=== handle ===` block per product, then import via the
Descriptions tab or:
```
python3 -c "
import app
products = app.load_products()
descs = app.parse_descriptions(open('<file>').read())
for p in products:
    if p['handle'] in descs:
        p['generated_description'] = descs[p['handle']]; p['description_approved'] = False
app.save_products(products)
"
```

## 6. Color Family metafield + Color Number tag (Neil's scheme)
Once products are live in Shopify (approved + imported), use the **"Compute Color
Family Proposals"** button on the Shopify tab -- it runs `classify_color_family()`
(Neil's digit map + the metallic/forced-exception/composite-lining/017xx rules, see
`margola_color_family_taxonomy` memory for the full writeup) against every approved
product, shows a review table, and only writes `custom.color_family` + the Color
Number tag for rows you confirm. Anything the classifier can't confidently place
comes back flagged "needs review" with a blank family for you to fill in by hand.
If a new category surfaces a color-number pattern the classifier doesn't recognize,
add it to `COLOR_FAMILY_KEYWORDS`/`COLOR_FAMILY_FORCED` in app.py rather than
guessing in the UI every time.

## 7. Before the real Shopify export
- Re-run the color-taxonomy check in step 3.
- Spot check `/api/export/shopify-test.csv` with 2-3 approved products first.
