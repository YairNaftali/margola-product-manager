# Margola Product Manager Roadmap

## Phase 1 — Stabilize Shopify Import
- [x] Import spreadsheet
- [x] Product review cards
- [x] Approve / skip workflow
- [x] Shopify CSV export
- [x] Variant model for Mini / Factory packs
- [x] Convert oz to grams for Shopify CSV
- [x] Disable inventory tracking
- [x] Upload images to Shopify Files
- [ ] Clean CSV exporter
- [ ] Confirm 1-product Shopify import works perfectly

## Phase 2 — Shopify GraphQL Sync
- [x] Confirm GraphQL connection works
- [x] Confirm product lookup by handle works
- [ ] Sync text metafields
- [ ] Sync Size metaobject reference
- [ ] Sync Color metaobject reference
- [ ] Sync Bead Shape metaobject reference
- [ ] Sync Bundle Pack Options metaobject reference

## Phase 3 — Architecture Cleanup
- [ ] Split app.py into modules
- [ ] Move Shopify API logic into shopify.py
- [ ] Move CSV export logic into export.py
- [ ] Move spreadsheet parsing into parser.py
- [ ] Move validation into validation.py
- [ ] Add safer error messages

## Phase 4 — UI Improvements
- [ ] Image status indicators
- [ ] Shopify sync status
- [ ] Export preflight checklist
- [ ] Missing data report
- [ ] Description review/import
