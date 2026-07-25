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
