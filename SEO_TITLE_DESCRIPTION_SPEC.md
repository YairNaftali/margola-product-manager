# SEO Title / Description — build spec

Written 2026-07-27 to hand off building this into the Product Manager app itself.
Currently this is NOT part of the CSV import pipeline or any app.py function —
every product's `seo.title` / `seo.description` on the live store has been set
by one-off scripts run outside the tool. Goal: make this a real feature so new
imports get correct SEO fields without a manual follow-up pass every time.

## Why this matters / history

On 2026-07-09 an earlier pass set `seo.title`/`seo.description` on 219/229
products. By 2026-07-23, re-auditing the full catalog (now 395 products) found
only 10 products had *any* value in those fields — and 8 of those 10 had the
**exact same title and description**, copy-pasted from one Rhinestone product
(the "8SS Chaton Rose... 10 Gross (1,440 pcs)" text) onto 8 different size
variants (8ss/10ss/12ss/16ss/20ss/30ss/34ss/40ss), each describing the wrong
size. Root cause was never confirmed (best guess: a loop variable that didn't
get reset between iterations in whatever script set these originally) — no
reimport happened since 7/9 that would otherwise explain it.

**Lesson baked into this spec:** whatever re-generates these fields must be
safe to re-run against the *whole* catalog periodically, not just newly
created products — don't assume a product having non-empty `seo.title` means
it's correct. See the flagging checklist below.

## SEO Title generation

Reuse `make_alt_text(title)` — already exists in `app.py`, already used for
the Image Alt Text CSV column. Do not write a second implementation.

```python
ALT_TEXT_PROTECTED_TERMS = {"AB": "AB", "DK": "Dark", "2XAB": "2X AB"}

def make_alt_text(title):
    t = clean(title)
    t = re.sub(r"\s*[–-]\s*\d+\s*Gross.*$", "", t, flags=re.I)   # strip "– N Gross" suffix
    t = re.sub(r"\s*\(\d[\d,]*\s*pcs\)\s*$", "", t, flags=re.I)  # strip "(N,NNN pcs)" suffix
    t = re.sub(r"\s+", " ", t).strip(" -–")
    if not t.isupper():
        return t
    t = re.sub(r"\bLT\.?\s*", "LIGHT ", t)
    t = re.sub(r"\s+", " ", t).strip()
    def repl(m):
        w = m.group(0)
        key = w.rstrip(".").upper()
        return ALT_TEXT_PROTECTED_TERMS.get(key, w.title())
    return re.sub(r"[A-Za-z0-9/.'-]+", repl, t)
```

**Critical Shopify quirk — must handle or writes silently no-op:** if
`seo.title` is set to a string byte-identical to the product's own `title`,
Shopify accepts the mutation (`userErrors: []`) but silently stores `null`
instead. Confirmed live 2026-07-23, hit 204 of 393 products in one run (any
product whose title needed no pack-suffix cleanup, so `make_alt_text(title)`
returned the title unchanged).

**Fix:** after generating the SEO title, compare it to the raw title. If
identical, append `" | Margola"`:

```python
seo_title = make_alt_text(title)
if seo_title == title:
    seo_title = f"{title} | Margola"
```

## SEO Description generation

Derive from the product's existing `descriptionHtml` — don't invent new copy
per product, just clean up and truncate what's already there.

```python
import re, html

def strip_html(raw):
    txt = re.sub(r"<[^>]+>", " ", raw or "")
    txt = html.unescape(txt)
    return re.sub(r"\s+", " ", txt).strip()

def fix_shouty_caps(text):
    """Title-cases any run of 3+ consecutive ALL-CAPS words (several
    description templates — Roller Beads, 2/3 Cut — have ALL-CAPS headings
    baked into the stored HTML itself, not just CSS text-transform)."""
    def repl(m):
        words = m.group(0).split()
        out = []
        for w in words:
            key = w.rstrip(".,").upper()
            out.append(ALT_TEXT_PROTECTED_TERMS.get(key, w.title()))
        return " ".join(out)
    return re.sub(r"\b(?:[A-Z][A-Z'/-]*\s+){2,}[A-Z][A-Z'/-]*\b", repl, text)

def truncate_words(txt, limit=160):
    """Cut at the last full sentence inside the limit if there is one far
    enough in to be worthwhile; otherwise cut at the last full word. Never
    cut mid-word — this was wrong in an early version of this script and
    produced descriptions ending on a dangling half-word."""
    if len(txt) <= limit:
        return txt
    window = txt[:limit]
    last_period = window.rfind(". ")
    if last_period > 80:
        return window[:last_period + 1]
    return window.rsplit(" ", 1)[0].rstrip(",.;:- ")

def gen_description(title, description_html):
    plain = strip_html(description_html)
    ct = make_alt_text(title)
    if not plain:
        # fallback template — only used when the product has no description at all
        return f"Shop {ct} from Margola — quality Czech glass beads & findings for jewelry, costume, and craft projects."
    plain = fix_shouty_caps(plain)
    if plain.lower().startswith(ct.lower()):
        plain = plain[len(ct):].strip(" -–")   # drop a leading repeat of the title
    return truncate_words(plain, 160)
```

## Second Shopify quirk — read-after-write lag

Immediately after a large batch of `productUpdate` mutations, a follow-up
GraphQL read can show stale/incomplete results for tens of seconds (saw a
clean 0-error push, then an audit right after reported 203 "still missing,"
then 0 missing on every re-check moments later). **Don't treat an audit run
immediately following a bulk push as authoritative** — if building an
automatic verify-after-write step, wait at least ~30-60s before concluding a
field didn't take, or re-check once before flagging it as failed.

## Recommended: build as an explicit action, with a review/flag queue

Don't wire this silently into the CSV import path with no human checkpoint —
model it on the existing "Add Color Family" button pattern (propose → review
table → confirm → apply), since the mechanical rules above can't catch every
edge case the way eyeballing a sample can. Flag these specific conditions for
review rather than auto-applying blind:

1. **Generated `seo_title == title` even after the `| Margola` fallback** —
   shouldn't happen, but worth asserting rather than assuming.
2. **`truncate_words` result ends without terminal punctuation and is exactly
   at the 160-char boundary** — a sign the cut landed awkwardly and a human
   should glance at it.
3. **`fix_shouty_caps` found and converted a run in this description** — not
   wrong, but worth a quick visual spot-check the first time a new category
   goes through this, in case the regex mis-cased something unexpected.
4. **Product had no `descriptionHtml` at all, so it's about to get the
   generic fallback sentence** — one or two of these is fine; if a whole
   batch/category is hitting the fallback, that's a missing-content gap
   worth surfacing, not a code bug to silently paper over.

## Operational note

Whatever this becomes (button, CLI script, pipeline step), it should be
**safe and normal to re-run against the entire catalog**, not just products
created since the last run — per the 7/9 incident above, "already has a
value" is not proof the value is currently correct.

## Implementation status (2026-07-25)

Built as specified: `strip_html`/`fix_shouty_caps`/`truncate_words`/
`gen_seo_title`/`gen_description` in `app.py`, reading the whole live catalog
via `shopify_list_products_for_seo()` (not local `products.json`), exposed as
`GET /api/shopify/seo-proposals` + `POST /api/shopify/apply-seo`, with a
review table on the Shopify tab mirroring the Color Family button pattern.
One deliberate deviation from a literal reading: flagged rows (any of the 4
conditions above) default to **unchecked** in the review table rather than
checked, since the spec says to flag for review rather than auto-apply blind.

**The "Second Shopify quirk" read-after-write verify step was deliberately
NOT built** (the spec marks it optional). `productUpdate`'s own `userErrors`
already tells you definitively whether a write was accepted, and that's
already surfaced in `shopify_apply_seo()`'s `errors` list — a delayed
re-query mainly protects against silent no-ops, and the specific silent-no-op
failure mode this spec describes (`seo.title == title`) is already prevented
structurally by the `| Margola` fallback + the `seo_title_still_matches_title`
flag, not something a 30-60s-later re-check would catch better.

**Update 2026-07-25, two follow-up fixes after first live use:**
1. **`strip_html()` bug:** replacing every HTML tag with a literal space put a
   stray space before punctuation whenever a closing inline tag (e.g. `</span>`
   from Google-Sheets-pasted descriptions) sat directly against it —
   `COATING</span>. Each` became `COATING . Each`. This was never a live data
   typo, purely a generator artifact (caught before ever calling `apply-seo`,
   so nothing live was wrong). Fixed by collapsing whitespace immediately
   before `.,;:!?` as a final step in `strip_html()`.
2. **Review-table checkbox defaults recalibrated:** a live full-catalog run
   showed `shouty_caps_fixed` firing on 375/458 products (82%) — every
   category using an ALL-CAPS `<h2>` template (2 Cut, 3 Cut, Roller, Fire
   Polished), while the only unflagged category (Rhinestone Flatback) already
   uses title-case. Spot-checked several categories and the transformation is
   correct everywhere, so continuing to require a manual checkbox click on
   82% of the catalog was pure busywork, not real safety. Rows whose *only*
   flag is `shouty_caps_fixed` are now pre-checked too (see
   `seoNeedsManualReview()` in `app.js`) — the other 3 flags (missing
   description, awkward truncation, anomalous title match) still default to
   unchecked, since those are genuinely rarer and worth a human glance.

**Update 2026-07-25, further UX pass (Yair: "so I don't have to check off as
many boxes"):** discovered `shouty_caps_fixed` fires on both changed *and*
already-correct rows (it only measures whether the raw HTML had caps to fix,
not whether the live value already reflects that fix from an older script
run) — of 375 flagged rows, 179 were already correct. Reworked the review
table: already-correct (`changed: false`) rows are now hidden by default
behind a "show N already-correct row(s) too" toggle, and the badge column
shows a plain green "unchanged" for those instead of the (misleadingly
alarming) shouty warning badge. The visible table is now just the rows that
actually need something written, all pre-checked unless genuinely flagged —
so the normal flow is Compute → Apply with no per-row clicking needed, unless
a rarer flag (missing description, awkward truncation, anomalous title) is
present.

**If asked to add it later** (e.g. "add an audit button for the SEO push"):
build it as a separate, standalone action — not wired into `apply-seo` itself
(a mandatory 30-60s wait after every apply click would make the button feel
hung). Concretely: a new `GET /api/shopify/seo-audit` endpoint that re-queries
`shopify_list_products_for_seo()` and reports any product whose live
`seo.title`/`seo.description` doesn't match what `gen_seo_title`/
`gen_description` would currently generate for it — meant to be run some time
*after* an apply pass, as a separate button/step, not chained automatically
after it.
