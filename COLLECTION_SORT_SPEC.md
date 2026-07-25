# Collection Sort Order — build spec

Written 2026-07-27 to hand off building this into the Product Manager app.
Currently this is NOT part of the app at all — every collection's sort order
has been set by one-off scripts run outside the tool (all ~54 collections
were sorted once, 2026-07-23; 2 Cut Beads was re-sorted 2026-07-25 after the
10/0 batch was added). Goal: a repeatable action in the app instead of a
fresh script each time a collection needs re-sorting.

## Why this matters

`sortOrder: MANUAL` on a Shopify collection is a **one-time snapshot, not a
live rule.** Once set, the order only changes when you explicitly push a new
order — any product added to the collection afterward (new color, new size,
whatever) just appends at the end, it does not get inserted into its correct
position automatically. This is why the 2 Cut Beads collection needed
re-sorting after the 10/0 batch was imported: the 11/0-only sort from 7/23
had no way to know 10/0 existed yet. Any category that keeps getting new
colors/sizes over time will need this repeated, indefinitely — that's the
whole reason to build it as a button instead of a script.

## Step 1: derive a sort key per product from its SKU

Don't parse the title — SKUs are the reliable structured source. Example SKU:
`2CUT-11/0-96090-mini` splits on `-` into `["2CUT", "11/0", "96090", "mini"]`.

```python
import re

# Only needed for collections that span more than one size (e.g. 2 Cut
# Beads' 10/0 + 11/0). Explicit map, not inferred — extend per category.
SIZE_ORDER = {"10/0": 0, "11/0": 1}

def sort_key(sku, color_name_fallback=""):
    parts = sku.split("-")
    size = parts[1] if len(parts) > 1 else ""
    color_raw = parts[2] if len(parts) > 2 else ""
    m = re.match(r"(\d+)", color_raw)
    if m:
        color_num = int(m.group(1))
        color_sort = (0, color_num, color_raw)          # numeric codes sort first, by number
    else:
        color_sort = (1, 0, color_name_fallback.upper()) # no numeric code -> alphabetical fallback
    return (SIZE_ORDER.get(size, 99), *color_sort)
```

**The numeric-vs-alphabetical fallback matters:** most product lines carry a
numeric color code in the SKU and sort correctly by that. A few (Chunky Mix,
Leather Cord) only have color *names*, no numeric code — those fall back to
plain alphabetical sort on the name instead of crashing or defaulting to 0.

**The size-grouping question is a judgment call, not something to infer.**
For 2 Cut Beads I asked whether the two sizes should be grouped (all 10/0
together, then all 11/0) or interleaved by color regardless of size —
grouping was chosen. If a collection ever has more than one size again,
**surface this as a choice in the UI**, don't default silently either way —
the two options produce genuinely different, equally defensible browsing
experiences.

## Step 2: apply the order in Shopify — two calls

```python
# 1. Collection must be in manual mode before a manual order can stick
mutation_set_manual = """
mutation SetManual($input: CollectionInput!) {
  collectionUpdate(input: $input) { userErrors { field message } }
}
"""
shopify_graphql(mutation_set_manual, {"input": {"id": collection_id, "sortOrder": "MANUAL"}})

# 2. Push the full order. sorted_products is the product list already
# ordered by sort_key() above; newPosition is just its 0-indexed rank.
mutation_reorder = """
mutation Reorder($id: ID!, $moves: [MoveInput!]!) {
  collectionReorderProducts(id: $id, moves: $moves) {
    job { id done }
    userErrors { field message }
  }
}
"""
moves = [{"id": p["id"], "newPosition": str(i)} for i, p in enumerate(sorted_products)]
result = shopify_graphql(mutation_reorder, {"id": collection_id, "moves": moves})
job_id = result["collectionReorderProducts"]["job"]["id"]
```

`collectionReorderProducts` runs as an **async job** — poll until it reports
done before considering the operation finished:

```python
job_query = '{ job(id: "%s") { done } }' % job_id
# poll job_query on a short interval until data["job"]["done"] is true
```

## Step 3: verify

Re-query the collection with `sortKey: COLLECTION_DEFAULT` and confirm the
live order actually matches what was pushed — don't just trust a clean
`userErrors: []` on the mutation:

```graphql
{
  collection(id: "...") {
    sortOrder
    products(first: 10, sortKey: COLLECTION_DEFAULT) { edges { node { title } } }
  }
}
```

## Recommended: build as a "Re-sort collection" action per category

Model it as an explicit button (like "Add Color Family"), not something that
runs automatically on import — it's cheap and safe to re-run (idempotent:
re-sorting into the same order is a no-op), so there's no harm in making it
available on demand whenever a category gets new products, rather than
trying to trigger it automatically at the end of an import.

If a category ever gets more than one size that needs grouping, expose the
group-vs-interleave choice as a prompt/setting rather than hardcoding it —
see the note under Step 1.
