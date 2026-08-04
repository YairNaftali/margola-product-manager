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
  it gets picked up automatically), `type` (see step 1b below), `shopify_category`,
  and **`sync`** -- the real
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

## 1b. Type (the "Type" storefront filter)
- `data/spreadsheet_types.json` -> `"type"` needs a short, clean value for the new
  category -- this feeds `custom.type` (a `list.single_line_text_field`, same
  serialization rules as `bead_shape` in step 2 below), which is what lets a broad
  collection containing multiple sub-lines show a "Type" filter (e.g. Czech Glass
  Beads splits into Roller/Crow via this field; Rhinestone Banding is tagged "Metal
  Set" even as the sole occupant of its collection, so a future sibling line drops
  in cleanly). Style: drop words the parent collection/category name already implies
  (`Metal Set`, not `Metal Set Rhinestone Banding`; `Roller`, not `Roller Beads`).
- If a single category actually contains two distinct sub-types by nature (not just
  shape) -- e.g. Cameos and Intaglios, which is one `spreadsheet_type` but two real
  product lines -- the JSON's `type` value is just the default; add a small
  per-row override in `infer()` next to the Roller Beads special case above (see
  the `cameos-intaglios` block: overrides to `"Intaglio"` when the raw Color Name
  contains that word, defaulting to `"Cameo"` otherwise). Don't confuse this with
  `bead_shape`'s `populate_bead_shape_from_shape` mechanism -- that pulls the row's
  own Shape column value directly, which doesn't apply here since Cameo/Intaglio
  isn't its own column.
- Backfilled live 2026-08-03 across the entire catalog (515 products, all
  categories including several with no `spreadsheet_types.json` entry at all --
  those were tagged directly via `metafieldsSet`, not through this pipeline, since
  they were never imported through this tool to begin with). See
  `margola_product_manager` memory for the full label table.

## 2. Bead shape taxonomy
- `data/shopify_taxonomy_map.json` -> `"bead_shape"` needs an entry for the new
  category: `{"label": "..."}` -- this is the literal text `resolver.bead_shape_label()`
  writes into the CSV's `custom.bead_shape` column, which is the field that
  actually powers the storefront Bead Shape filter (see
  `margola_shopify_filter_architecture` memory). The old `shopify.bead-shape`
  metaobject-reference field/CSV column was removed 2026-07-25 -- it never
  showed custom labels on the storefront and wasn't feeding anything else
  useful, so don't reintroduce it.
- **The CSV cell must contain the plain label text only (e.g. `Bugle`), never
  `json.dumps([label])`.** `custom.bead_shape` is a `list.single_line_text_field`,
  and its *final live stored value* does look like `["Bugle"]` -- but that's
  Shopify's CSV importer wrapping your plain cell text into a list itself.
  Pre-encoding it yourself double-wraps it into `["[\"Bugle\"]"]` (a real bug,
  shipped and caught 2026-07-25 onboarding Bugle Beads -- see
  `margola_shopify_filter_architecture` Fact 6). Only a direct GraphQL
  `metafieldsSet`/`metaobjectCreate` call wants the JSON-encoded array string;
  CSV import never does.

## 2b. Bead size taxonomy
- `data/shopify_taxonomy_map.json` -> `"size"` needs an entry per distinct size this
  category uses: `{"label": "...", "base_gid": "gid://shopify/TaxonomyValue/..."}`.
  `resolver.size_handle()` resolves/creates the `shopify--size` metaobject and
  writes its handle into `shopify.size` (this one *is* a working metaobject-reference
  filter, unlike bead shape/color -- Shopify's Size category apparently doesn't
  collapse custom labels the same way).
- Check for an existing live metaobject before adding a new `base_gid` -- reuse it
  exactly if the size already exists:
  `python3 -c "import app; [print(n['displayName'], n['handle']) for n in app.shopify_list_metaobjects('shopify--size')]"`
  then fetch its `taxonomy_reference` field via a `metaobjectByHandle` query. Bead
  sizes in the `X/0 - Y.Ymm` format (10/0, 11/0, 9/0, 12/0, etc.) all share
  `gid://shopify/TaxonomyValue/2878` -- confirmed live 2026-07-25 across four
  existing entries. Only look up a different GID for a genuinely different size
  format (mm-only beads, stone sizes in ss, etc. use different base values).

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

## 8. Collection Sort Order
Once the category's products are live in their target collection(s), use the
**"Collection Sort Order"** box on the Shopify tab (enter the collection
handle, e.g. `2-cut-beads`) -- computes a proposed order from each product's
SKU (color code, numeric-first with alphabetical fallback for lines that only
have color names, e.g. Chunky Mix/Leather Cord) and shows a before/after
review table before pushing. Only needed once per collection whenever new
products get added to it -- a manual sort is a one-time snapshot, not a live
rule, so new products just append at the end until this is re-run. If the
category shares a collection with another size (like 2 Cut Beads' 10/0 +
11/0), add an entry to `COLLECTION_SIZE_ORDER` in app.py and the UI will
surface a group-by-size vs. interleave-by-color choice automatically -- see
`COLLECTION_SORT_SPEC.md` for the full build spec.

## 9. SEO Title/Description
Nothing category-specific to configure here -- the **"SEO Title + Description"**
box on the Shopify tab audits the *entire* live catalog every time (not just
this category), deriving from each product's own title/description, so a new
category's products just get picked up automatically the next time it's run.
See `SEO_TITLE_DESCRIPTION_SPEC.md` for the full build spec and its
"Implementation status" section for known quirks/decisions.
