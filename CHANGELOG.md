# Changelog

## Unreleased

### Added
- Shopify client-credentials connection flow.
- Shopify Files upload testing.
- Variant model for pack options.
- Shopify CSV test export.
- GraphQL product lookup test.

### Changed
- Inventory export should not track inventory.
- Variant weights are converted from ounces to grams for Shopify CSV.
- Metafield-heavy CSV import is being replaced with GraphQL sync.

### Fixed
- Shopify token flow after app installation.
- Python SSL certificate issues using certifi.
- False SKU assumptions for missing Mini/Factory packs.

### Known Issues
- CSV exporter still needs cleanup.
- Metaobject fields cannot be imported as plain CSV text.
- Image filename matching needs improvement.
