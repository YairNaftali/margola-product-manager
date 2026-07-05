import csv, io, json, os, re, uuid, time, mimetypes, urllib.parse, urllib.request, ssl, certifi
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse
import openpyxl

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(APP_DIR, "data")
UPLOAD_DIR = os.path.join(APP_DIR, "uploads")
PRODUCTS_PATH = os.path.join(DATA_DIR, "products.json")
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(UPLOAD_DIR, exist_ok=True)

def clean(v):
    if v is None: return ""
    if isinstance(v, float) and v.is_integer(): return str(int(v))
    return str(v).strip()

def slugify(v):
    v = clean(v).lower()
    v = re.sub(r"[^a-z0-9]+", "-", v)
    return re.sub(r"-+", "-", v).strip("-")

def normalize_header(h):
    return re.sub(r"[^a-z0-9]+", "_", clean(h).lower()).strip("_")

def money(v):
    if v in (None, ""): return ""
    try: return f"{float(v):.2f}"
    except Exception: return ""

def weight_oz(v):
    text = clean(v).lower().replace("ounces", "oz").replace("ounce", "oz")
    m = re.search(r"[-+]?\d*\.?\d+", text)
    return m.group(0) if m else ""

def get_first(row, names):
    for n in names:
        k = normalize_header(n)
        if k in row and row[k] not in (None, ""):
            return row[k]
    return ""

def row_dict(headers, values):
    return {normalize_header(headers[i]): values[i] if i < len(values) else None for i in range(len(headers))}

def infer(sheet_name, color_name, source_filename=""):
    s = sheet_name.lower()
    fn = source_filename.lower()
    c = clean(color_name).lower()
    BEADS_CATEGORY = "Arts & Entertainment > Hobbies & Creative Arts > Arts & Crafts > Art & Crafting Materials > Embellishments & Trims > Beads"
    if "crow" in s:
        return {"collection":"Czech Glass Beads","title_prefix":"Czech Glass","subcategory":"Crow Beads","bead_shape":"Crow Beads","color_type":"Opaque","factory_qty_standard":"1000 pieces","mini_qty_standard":"100 pieces","populate_bead_shape_from_shape":False,"shopify_category":BEADS_CATEGORY}
    if "roller" in s:
        if "transparent" in s or "transparent" in c or "transpaent" in c: ct = "Transparent"
        elif "opaque" in s or "opaque" in c: ct = "Opaque"
        else: ct = ""
        return {"collection":"Czech Glass Beads","title_prefix":"Czech Glass","subcategory":"Roller Beads","bead_shape":"Roller Beads","color_type":ct,"factory_qty_standard":"1200 beads","mini_qty_standard":"144 beads","populate_bead_shape_from_shape":False,"shopify_category":BEADS_CATEGORY}
    if "stringing" in fn or "leather cord" in s:
        return {"collection":"Tools & Stringing Materials","title_prefix":"","subcategory":"Leather Cord","bead_shape":"","color_type":"","factory_qty_standard":"","mini_qty_standard":"","populate_bead_shape_from_shape":False,"shopify_category":"Arts & Entertainment > Hobbies & Creative Arts > Arts & Crafts > Art & Crafting Materials > Crafting Fibers > Jewelry & Beading Cord"}
    if "lalique" in fn or "plexy" in fn:
        return {"collection":"Plexi-Lalique Flowers & Leaves","title_prefix":"Plexi-Lalique","subcategory":"","bead_shape":"","color_type":"","factory_qty_standard":"","mini_qty_standard":"","populate_bead_shape_from_shape":True,"shopify_category":BEADS_CATEGORY}
    return {"collection":"","title_prefix":"","subcategory":"","bead_shape":"","color_type":"","factory_qty_standard":"","mini_qty_standard":"","populate_bead_shape_from_shape":False,"shopify_category":""}

def title_for(prefix, descriptor, size, color_name):
    # No assumptions: do not invent manufacturer/brand.
    return " ".join([x for x in [clean(prefix), clean(descriptor), clean(size), clean(color_name)] if x])

def image_for(factory_style, color_name, subcategory, size):
    base = clean(factory_style) or " ".join([subcategory, size, color_name])
    return f"{slugify(base)}.jpg" if base else ""


def build_variants(factory_style, factory_price, factory_weight, factory_qty_desc, mini_style, mini_price, mini_weight, mini_qty):
    variants = []

    # No assumptions:
    # A pack variant only exists if it has its own SKU AND at least some real pack data.
    # Do not copy Mini into Factory or Factory into Mini.
    if clean(mini_style) and (clean(mini_price) or clean(mini_weight) or clean(mini_qty)):
        variants.append({
            "name": "Mini Pack",
            "option_value": "mini-pack",
            "sku": clean(mini_style),
            "price": money(mini_price),
            "weight_oz": weight_oz(mini_weight),
            "quantity": clean(mini_qty),
        })

    if clean(factory_style) and (clean(factory_price) or clean(factory_weight) or clean(factory_qty_desc)):
        variants.append({
            "name": "Factory Pack",
            "option_value": "factory-pack",
            "sku": clean(factory_style),
            "price": money(factory_price),
            "weight_oz": weight_oz(factory_weight),
            "quantity": clean(factory_qty_desc),
        })

    return variants

def validate(p):
    variants = p.get("variants", [])

    checks = {
        "Title": bool(p.get("title")),
        "Handle": bool(p.get("handle")),
        "At least one variant": bool(variants),
        "All variant SKUs": bool(variants) and all(v.get("sku") for v in variants),
        "All variant prices": bool(variants) and all(v.get("price") for v in variants),
        "All variant weights": bool(variants) and all(v.get("weight_oz") for v in variants),
        "Collection": bool(p.get("collection")),
        "Subcategory": bool(p.get("subcategory")),
        "Color type": bool(p.get("color_type")),
        "Bead shape": bool(p.get("bead_shape")),
        "Image filename": bool(p.get("image_filename")),
        "Description": bool(p.get("generated_description") or p.get("source_description")),
    }

    score = int(round(sum(1 for ok in checks.values() if ok) / len(checks) * 100))
    warnings = [k for k, ok in checks.items() if not ok]

    if p.get("notes"):
        warnings.append(p["notes"])

    return {"checks": checks, "score": score, "warnings": warnings}

def parse_xlsx(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    products = []
    for ws in wb.worksheets:
        headers = [c.value for c in ws[1]]
        if not any(headers): continue
        for row_num, vals in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not any(vals): continue
            row = row_dict(headers, vals)
            factory_style = get_first(row, ["FACTORY PACK STYLE #","FACTORY PACK       STYLE #","FACTORY PACK STYLE"])
            if not clean(factory_style): continue
            size = get_first(row, ["BEAD SIZE DIAMETER  MM","BEAD SIZE","SIZE","Milimeter size"])
            color_number = get_first(row, ["COLOR NUMBER"])
            color_name = get_first(row, ["COLOR NAME"])
            shape = get_first(row, ["SHAPE"])
            source_description = get_first(row, ["DESCRIPTION"])
            factory_qty = get_first(row, ["FACTORY PACK UNIT QUANTITY"])
            factory_qty_desc = get_first(row, ["UNIT QUANTITY DESCRIPTION","UNIT QUANTITY DESCRIPTION FACTORY PACK"])
            factory_price = get_first(row, ["UNIT PRICE PER FACTORY PACK","UNIT PRICE PER BAG"])
            factory_weight = get_first(row, ["WEIGHT PER FACTORY PACK"])
            mini_style = get_first(row, ["MINI PACK STYLE #","MINI PACK       STYLE #"])
            mini_qty = get_first(row, ["MINI PACK QUANTITY","MINI PACK QUANTITY DESCRIPTION"])
            mini_price = get_first(row, ["UNIT PRICE PER MINI PACK"])
            mini_weight = get_first(row, ["WEIGHT PER MINI PACK"])
            inf = infer(ws.title, color_name, os.path.basename(path))
            if not inf["collection"]:
                raise ValueError(f"Could not detect a known product category for {os.path.basename(path)!r} (sheet {ws.title!r}, row {row_num}). Add a matching rule to infer() before importing this file.")
            title_descriptor = clean(shape) or inf["subcategory"]
            title = title_for(inf["title_prefix"], title_descriptor, size, color_name)
            bead_shape_value = clean(shape) if inf["populate_bead_shape_from_shape"] and clean(shape) else inf["bead_shape"]
            row_text = " ".join(clean(x) for x in vals).lower()
            notes = []
            if "missing phot" in row_text or "mising phot" in row_text: notes.append("Source note: missing photo")
            if clean(factory_price) and not money(factory_price): notes.append(f"Invalid factory price in source: {clean(factory_price)!r}")
            if clean(mini_price) and not money(mini_price): notes.append(f"Invalid mini price in source: {clean(mini_price)!r}")
            if clean(mini_style) and clean(size) and clean(size).lower() not in clean(mini_style).lower():
                notes.append(f"Mini pack style # may not match this row's size ({clean(size)!r}): {clean(mini_style)!r}")
            p = {
                "id":str(uuid.uuid4()), "source_file":os.path.basename(path), "source_sheet":ws.title, "source_row":row_num,
                "approved":False, "skipped":False, "status":"Needs Review",
                "title":title, "handle":slugify(title), "brand":"", "vendor":"Margola",
                "collection":inf["collection"], "subcategory":inf["subcategory"], "color_type":inf["color_type"], "bead_shape":bead_shape_value,
                "shopify_category":inf["shopify_category"],
                "size":clean(size), "color_number":clean(color_number), "color_name":clean(color_name),
                "image_filename":image_for(factory_style,color_name,title_descriptor,size), "image_src":"", "image_alt":title,
                "factory_style":clean(factory_style), "factory_quantity":clean(factory_qty),
                "factory_quantity_description":clean(factory_qty_desc) or inf["factory_qty_standard"],
                "factory_price":money(factory_price), "factory_weight_oz":weight_oz(factory_weight),
                "mini_style":clean(mini_style), "mini_quantity":clean(mini_qty) or inf["mini_qty_standard"],
                "mini_price":money(mini_price), "mini_weight_oz":weight_oz(mini_weight),

                "variants": build_variants(
                    factory_style=factory_style,
                    factory_price=factory_price,
                    factory_weight=factory_weight,
                    factory_qty_desc=clean(factory_qty_desc) or inf["factory_qty_standard"],
                    mini_style=mini_style,
                    mini_price=mini_price,
                    mini_weight=mini_weight,
                    mini_qty=clean(mini_qty) or inf["mini_qty_standard"],
                ),

                "source_description":clean(source_description), "generated_description":"", "description_approved":False,
                "notes":"; ".join(notes),
            }
            products.append(p)
    handles, skus = {}, {}
    for p in products:
        handles.setdefault(p["handle"], []).append(p["id"])
        for v in p.get("variants", []):
            sku = v.get("sku")
            if sku:
                skus.setdefault(sku, []).append(p["id"])
    dup_handles = {k for k,v in handles.items() if len(v)>1}
    dup_skus = {k for k,v in skus.items() if len(v)>1}
    for p in products:
        more = []
        if p["handle"] in dup_handles: more.append("Duplicate handle")
        for v in p.get("variants", []):
            if v.get("sku") in dup_skus:
                more.append(f"Duplicate variant SKU: {v.get('sku')}")
        if more: p["notes"] = "; ".join([x for x in [p["notes"]] + more if x])
        p["validation"] = validate(p)
        p["status"] = "Needs Review" if p["validation"]["warnings"] else "Ready for Review"
    return products

def load_products():
    if not os.path.exists(PRODUCTS_PATH): return []
    with open(PRODUCTS_PATH, "r", encoding="utf-8") as f: return json.load(f)

def save_products(products):
    for p in products: p["validation"] = validate(p)
    with open(PRODUCTS_PATH, "w", encoding="utf-8") as f: json.dump(products, f, indent=2)

def summary(products):
    return {"total":len(products),"approved":sum(p.get("approved") for p in products),"skipped":sum(p.get("skipped") for p in products),"needs_review":sum(1 for p in products if not p.get("approved") and not p.get("skipped")),"with_warnings":sum(1 for p in products if p.get("validation",{}).get("warnings")),"ready_100":sum(1 for p in products if p.get("validation",{}).get("score")==100)}

def detected_categories(products):
    counts = {}
    for p in products:
        key = (p.get("collection",""), p.get("subcategory",""))
        counts[key] = counts.get(key, 0) + 1
    return [{"collection":k[0],"subcategory":k[1],"count":v} for k,v in sorted(counts.items())]

def review_csv(products):
    fields = ["status","approved","skipped","validation_score","warnings","title","handle","collection","subcategory","color_type","bead_shape","size","color_number","color_name","image_filename","image_src","factory_style","factory_price","factory_weight_oz","mini_style","mini_price","mini_weight_oz","source_sheet","source_row"]
    out = io.StringIO(); w = csv.DictWriter(out, fieldnames=fields); w.writeheader()
    for p in products:
        v = p.get("validation",{})
        row = {f:p.get(f,"") for f in fields}; row["validation_score"] = v.get("score",""); row["warnings"] = "; ".join(v.get("warnings",[]))
        w.writerow(row)
    return out.getvalue()


def oz_to_grams(value):
    try:
        return str(round(float(value) * 28.3495, 2))
    except:
        return ""

def shopify_rows(products, approved_only=True, limit=None, resolver=None):
    if resolver is None:
        resolver = MetaobjectResolver()
    rows, count = [], 0

    for p in products:
        if approved_only and (not p.get("approved") or p.get("skipped")):
            continue

        variants = p.get("variants", [])

        # Backward compatibility for old imported data without variants yet.
        if not variants:
            variants = build_variants(
                factory_style=p.get("factory_style"),
                factory_price=p.get("factory_price"),
                factory_weight=p.get("factory_weight_oz"),
                factory_qty_desc=p.get("factory_quantity_description"),
                mini_style=p.get("mini_style"),
                mini_price=p.get("mini_price"),
                mini_weight=p.get("mini_weight_oz"),
                mini_qty=p.get("mini_quantity"),
            )

        if not variants:
            continue

        if limit is not None and count >= limit:
            break

        count += 1
        body = p.get("generated_description") or p.get("source_description") or ""

        common = {
            "Handle": p["handle"],
            "Vendor": p.get("vendor") or "Margola",
            "Product Category": p.get("shopify_category", ""),
            "Type": p["subcategory"] or p.get("collection", ""),
            "Tags": ", ".join(x for x in [p["collection"], p["subcategory"], p["color_type"], p["bead_shape"], p["size"]] if x),
            "Published": "TRUE",
            "Variant Inventory Tracker": "",
            "Variant Inventory Qty": "",
            "Variant Inventory Policy": "deny",
            "Variant Fulfillment Service": "manual",
            "Variant Requires Shipping": "TRUE",
            "Variant Taxable": "TRUE",
            "Variant Weight Unit": "oz",
            "Status": "active"
        }

        for idx, variant in enumerate(variants):
            row = dict(common)
            row.update({
                "Title": p["title"] if idx == 0 else "",
                "Body (HTML)": body if idx == 0 else "",
                "Option1 Name": "Bundle Pack Options" if idx == 0 else "",
                "Option1 Value": variant.get("option_value", ""),
                "Option1 Linked To": "product.metafields.custom.bundle_pack_options" if idx == 0 else "",
                "Bundle Pack Options (product.metafields.custom.bundle_pack_options)": "; ".join(v["option_value"] for v in variants) if idx == 0 else "",
                "Color (product.metafields.shopify.color-pattern)": resolver.color_handle(p.get("color_name")) if idx == 0 else "",
                "Bead shape (product.metafields.shopify.bead-shape)": resolver.bead_shape_handle(p.get("bead_shape")) if idx == 0 else "",
                "Color Type (product.metafields.custom.color_type)": clean(p.get("color_type")) if idx == 0 else "",
                "Bead Size (product.metafields.custom.bead_size_mm)": clean(p.get("size")) if idx == 0 else "",
                "Variant SKU": variant.get("sku", ""),
                "Variant Price": variant.get("price", ""),
                "Variant Grams": oz_to_grams(variant.get("weight_oz", "")),
                "Image Src": p.get("image_src", "") if idx == 0 else "",
                "Image Alt Text": (p.get("image_alt") or p["title"]) if idx == 0 and p.get("image_src") else "",
            })
            rows.append(row)

    return rows

def shopify_csv(products, approved_only=True, limit=None):
    fields = [
        "Handle","Title","Body (HTML)","Vendor","Product Category","Type","Tags","Published",
        "Option1 Name","Option1 Value","Option1 Linked To",
        "Variant SKU","Variant Price","Variant Grams","Variant Weight Unit",
        "Variant Inventory Tracker","Variant Inventory Qty","Variant Inventory Policy",
        "Variant Fulfillment Service","Variant Requires Shipping","Variant Taxable",
        "Image Src","Image Alt Text",
        "Bundle Pack Options (product.metafields.custom.bundle_pack_options)",
        "Color (product.metafields.shopify.color-pattern)",
        "Bead shape (product.metafields.shopify.bead-shape)",
        "Color Type (product.metafields.custom.color_type)",
        "Bead Size (product.metafields.custom.bead_size_mm)",
        "Status"
    ]

    out = io.StringIO()
    w = csv.DictWriter(out, fieldnames=fields, extrasaction="ignore")
    w.writeheader()

    resolver = MetaobjectResolver()
    for row in shopify_rows(products, approved_only, limit, resolver=resolver):
        w.writerow(row)

    return out.getvalue()

def parse_descriptions(text):
    blocks, cur, lines = {}, None, []
    for line in text.splitlines():
        m = re.match(r"^\s*===\s*(.*?)\s*===\s*$", line)
        if m:
            if cur: blocks[cur] = "\n".join(lines).strip()
            cur, lines = m.group(1).strip(), []
        else: lines.append(line)
    if cur: blocks[cur] = "\n".join(lines).strip()
    return {k:v for k,v in blocks.items() if k and v}

def guess(path):
    if path.endswith(".css"): return "text/css; charset=utf-8"
    if path.endswith(".js"): return "application/javascript; charset=utf-8"
    if path.endswith(".html"): return "text/html; charset=utf-8"
    return "application/octet-stream"


# --- Shopify client-credentials integration ---
def load_env():
    env_path = os.path.join(APP_DIR, ".env")
    if os.path.exists(env_path):
        for line in open(env_path, encoding="utf-8"):
            line=line.strip()
            if line and not line.startswith("#") and "=" in line:
                k,v=line.split("=",1)
                os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))
load_env()
SHOPIFY_TOKEN_CACHE={"token":"","expires_at":0}
def shopify_config():
    return {"store":os.environ.get("SHOPIFY_STORE","").strip().replace("https://","").replace("http://","").strip("/"),"client_id":os.environ.get("SHOPIFY_CLIENT_ID","").strip(),"client_secret":os.environ.get("SHOPIFY_CLIENT_SECRET","").strip()}
def shopify_get_token(force=False):
    cfg=shopify_config(); now=time.time()
    if not cfg["store"] or not cfg["client_id"] or not cfg["client_secret"]: raise RuntimeError("Missing SHOPIFY_STORE, SHOPIFY_CLIENT_ID, or SHOPIFY_CLIENT_SECRET in .env")
    if not force and SHOPIFY_TOKEN_CACHE["token"] and SHOPIFY_TOKEN_CACHE["expires_at"]>now+60: return SHOPIFY_TOKEN_CACHE["token"]
    url=f"https://{cfg['store']}/admin/oauth/access_token"
    data=urllib.parse.urlencode({"grant_type":"client_credentials","client_id":cfg["client_id"],"client_secret":cfg["client_secret"]}).encode()
    req=urllib.request.Request(url,data=data,headers={"Content-Type":"application/x-www-form-urlencoded"},method="POST")
    with urllib.request.urlopen(req, timeout=30, context=ssl.create_default_context(cafile=certifi.where())) as resp: payload=json.loads(resp.read().decode())
    token=payload.get("access_token")
    if not token: raise RuntimeError(f"No access_token returned: {payload}")
    SHOPIFY_TOKEN_CACHE["token"]=token; SHOPIFY_TOKEN_CACHE["expires_at"]=now+int(payload.get("expires_in",86399)); return token
def shopify_graphql(query, variables=None):
    cfg=shopify_config(); token=shopify_get_token(); url=f"https://{cfg['store']}/admin/api/2026-04/graphql.json"
    body=json.dumps({"query":query,"variables":variables or {}}).encode()
    req=urllib.request.Request(url,data=body,headers={"Content-Type":"application/json","X-Shopify-Access-Token":token},method="POST")
    with urllib.request.urlopen(req, timeout=60, context=ssl.create_default_context(cafile=certifi.where())) as resp: payload=json.loads(resp.read().decode())
    if payload.get("errors"): raise RuntimeError(json.dumps(payload["errors"]))
    return payload["data"]
def shopify_test_connection(): return shopify_graphql("{ shop { name myshopifyDomain } }")["shop"]
def shopify_list_files(first=100):
    q="""query GetFiles($first:Int!){ files(first:$first){ edges{ node{ id alt createdAt ... on MediaImage{ image{ url } } ... on GenericFile{ url } } } } }"""
    data=shopify_graphql(q,{"first":first}); out=[]
    for e in data["files"]["edges"]:
        n=e["node"]; url=""
        if n.get("image"): url=n["image"].get("url") or ""
        if not url: url=n.get("url") or ""
        fn=urllib.parse.unquote(url.split("?")[0].rstrip("/").split("/")[-1]) if url else ""
        out.append({"id":n.get("id",""),"url":url,"filename":fn,"alt":n.get("alt",""),"createdAt":n.get("createdAt","")})
    return out
def file_map_path(): return os.path.join(DATA_DIR,"shopify_files.json")
def save_file_map(files): open(file_map_path(),"w",encoding="utf-8").write(json.dumps(files,indent=2))
def load_file_map():
    return json.load(open(file_map_path(),encoding="utf-8")) if os.path.exists(file_map_path()) else []
def apply_file_matches_to_products():
    products=load_products(); files=load_file_map(); by={f.get("filename","").lower():f.get("url") for f in files if f.get("filename") and f.get("url")}; matched=0
    for p in products:
        key=(p.get("image_filename") or "").lower()
        if key in by:
            p["image_src"]=by[key]; matched+=1
        elif p.get("bead_shape")=="Roller Beads":
            # Roller Beads only shoot one photo per color and reuse it for both
            # the 6mm and 9mm listing, so fall back to the other size's filename.
            alt_key=""
            if "roller-6mm-" in key: alt_key=key.replace("roller-6mm-","roller-9mm-")
            elif "roller-9mm-" in key: alt_key=key.replace("roller-9mm-","roller-6mm-")
            if alt_key and alt_key in by: p["image_src"]=by[alt_key]; matched+=1
    save_products(products); return matched,len(products)

DRIVE_FILELISTS = [
    os.path.join(os.path.expanduser("~"), "Downloads", "harddrive_filelist.txt"),
    os.path.join(os.path.expanduser("~"), "Downloads", "jumbo_filelist.txt"),
]
JUNK_PATH_MARKERS = ("._", ".DS_Store", ".psd", " - Copy", "---Copy")

def load_drive_index():
    # Reads the pre-generated drive filelist dumps (tab-separated: size, mtime, path)
    # rather than scanning the drives live, since they aren't always plugged in.
    roller_9mm, roller_6mm, crow, leather_cord, generic = {}, {}, {}, {}, {}
    found_any = False
    for list_path in DRIVE_FILELISTS:
        if not os.path.exists(list_path): continue
        found_any = True
        for line in open(list_path, encoding="utf-8", errors="replace").read().splitlines():
            parts = line.split("\t")
            if len(parts) < 3: continue
            full = parts[2]
            if any(m in full for m in JUNK_PATH_MARKERS): continue
            base = full.rsplit("/", 1)[-1].lower()
            fl = full.lower()
            generic.setdefault(base, full)
            if "/roller beads/9mm/" in fl: roller_9mm.setdefault(base, full)
            elif "/roller beads/6mm/" in fl: roller_6mm.setdefault(base, full)
            elif "/crow" in fl: crow.setdefault(base, full)
            elif "/leather cord/" in fl: leather_cord.setdefault(base, full)
    roller_union = dict(roller_6mm)
    for k, v in roller_9mm.items(): roller_union.setdefault(k, v)
    return {"found_any": found_any, "filelists": DRIVE_FILELISTS,
            "roller_9mm": roller_9mm, "roller_6mm": roller_6mm, "roller_union": roller_union,
            "crow": crow, "leather_cord": leather_cord, "generic": generic}

def _roller_candidates(color_number):
    cn = clean(color_number)
    cands = []
    m = re.search(r"use\s+(?:color\s+)?(?:the\s+)?([a-z0-9 .]+?)(?:\s+image\)?|\)|$)", cn, re.I)
    if m:
        hint = m.group(1).strip().lower().replace("metallc", "metallic")
        cands.append(hint.replace(" ", "-"))
        cands.append(re.sub(r"[^0-9a-z]", "", hint))
    raw = cn.split("(")[0].strip()
    is_matte = "matte" in raw.lower()
    digits = re.sub(r"[^0-9]", "", raw)
    if digits:
        variants = [digits]
        if len(digits) == 5 and digits.endswith("0"): variants.append(digits[:-1])
        for v in list(variants):
            cands.append(v)
            if is_matte: cands.append(v + "m")
    return [c for c in cands if c]

def _find_prefixed(pool, prefixes):
    for prefix in prefixes:
        for k, v in pool.items():
            if k.startswith(prefix): return v
    return None

def _find_suffixed(pool, suffixes):
    for suffix in suffixes:
        for k, v in pool.items():
            if k.endswith(suffix): return v
    return None

def _find_roller_photo(pool, cands):
    for c in cands:
        for key in (f"rb-{c}.jpg", f"rb{c}.jpg", f"rb-{c}.jpeg", f"rb{c}.jpeg"):
            if key in pool: return pool[key]
    for c in cands:
        found = _find_prefixed(pool, (f"rb-{c}-", f"rb-{c} ", f"rb{c}-", f"rb-{c}.", f"rb{c}."))
        if found: return found
    # Some files put the color code at the end instead, e.g.
    # "RB-gold-metallic-01710.jpg" for candidate "01710".
    for c in cands:
        if len(c) >= 4 and c.isalnum():
            found = _find_suffixed(pool, (f"-{c}.jpg", f"-{c}.jpeg"))
            if found: return found
    return None

def _leather_cord_candidate(image_filename):
    m = re.match(r"lc-(\d+)-(\d+)mm-(.+)\.jpg$", image_filename or "")
    if not m: return None
    whole, frac, color = m.groups()
    size = whole if frac == "0" else f"{whole}.{frac}"
    return size, color.replace("-", " ")

def resolve_photo_from_drive(product, index):
    # Read-only lookup: finds a real file on the indexed drives for a product
    # missing image_src. Does not touch Shopify or the filesystem.
    target = clean(product.get("image_filename"))
    if not target: return None
    key = target.lower()
    if product.get("bead_shape") == "Crow Beads":
        code = re.sub(r"[^0-9]", "", product.get("color_number") or "")
        found = _find_prefixed(index["crow"], (f"crow-9mm-{code}-", f"crow-9mm-{code} "))
        if found: return {"source_path": found, "target_filename": target, "reused_other_size": False}
    elif product.get("bead_shape") == "Roller Beads":
        is_9mm = "9mm" in key
        own_pool = index["roller_9mm"] if is_9mm else index["roller_6mm"]
        cands = _roller_candidates(product.get("color_number") or "")
        found = _find_roller_photo(own_pool, cands)
        if found: return {"source_path": found, "target_filename": target, "reused_other_size": False}
        found = _find_roller_photo(index["roller_union"], cands)
        if found: return {"source_path": found, "target_filename": target, "reused_other_size": True}
    elif product.get("subcategory") == "Leather Cord":
        parsed = _leather_cord_candidate(target)
        if parsed:
            size, color = parsed
            for k, v in index["leather_cord"].items():
                kk = k.replace("-", " ")
                if f"{size}mm" in kk and color in kk:
                    return {"source_path": v, "target_filename": target, "reused_other_size": False}
    if key in index["generic"]:
        return {"source_path": index["generic"][key], "target_filename": target, "reused_other_size": False}
    return None

def multipart_form_data(fields, files):
    boundary="----MargolaBoundary"+uuid.uuid4().hex; chunks=[]
    for k,v in fields.items(): chunks += [f"--{boundary}\r\n".encode(), f'Content-Disposition: form-data; name="{k}"\r\n\r\n'.encode(), str(v).encode(), b"\r\n"]
    for k,(filename,ctype,data) in files.items(): chunks += [f"--{boundary}\r\n".encode(), f'Content-Disposition: form-data; name="{k}"; filename="{filename}"\r\n'.encode(), f"Content-Type: {ctype}\r\n\r\n".encode(), data, b"\r\n"]
    chunks.append(f"--{boundary}--\r\n".encode()); return boundary,b"".join(chunks)
def shopify_upload_file(filepath, alt="", filename=None):
    filename=filename or os.path.basename(filepath); ctype=mimetypes.guess_type(filepath)[0] or "image/jpeg"; size=str(os.path.getsize(filepath))
    q1="""mutation stagedUploadsCreate($input:[StagedUploadInput!]!){ stagedUploadsCreate(input:$input){ stagedTargets{ url resourceUrl parameters{ name value } } userErrors{ field message } } }"""
    d=shopify_graphql(q1,{"input":[{"filename":filename,"mimeType":ctype,"httpMethod":"POST","resource":"FILE","fileSize":size}]})["stagedUploadsCreate"]
    if d.get("userErrors"): raise RuntimeError(json.dumps(d["userErrors"]))
    t=d["stagedTargets"][0]; fields={p["name"]:p["value"] for p in t["parameters"]}; data=open(filepath,"rb").read(); boundary,body=multipart_form_data(fields,{"file":(filename,ctype,data)})
    req=urllib.request.Request(t["url"],data=body,headers={"Content-Type":f"multipart/form-data; boundary={boundary}"},method="POST")
    urllib.request.urlopen(req, timeout=120, context=ssl.create_default_context(cafile=certifi.where())).read()
    q2="""mutation fileCreate($files:[FileCreateInput!]!){ fileCreate(files:$files){ files{ id alt fileStatus createdAt ... on MediaImage{ image{ url } } ... on GenericFile{ url } } userErrors{ field message } } }"""
    c=shopify_graphql(q2,{"files":[{"alt":alt or filename,"contentType":"IMAGE","originalSource":t["resourceUrl"]}]})["fileCreate"]
    if c.get("userErrors"): raise RuntimeError(json.dumps(c["userErrors"]))
    f=c["files"][0]; url=(f.get("image") or {}).get("url") or f.get("url") or ""
    return {"id":f.get("id",""),"filename":filename,"url":url,"alt":f.get("alt",""),"status":f.get("fileStatus",""),"createdAt":f.get("createdAt","")}

def shopify_find_or_create_collection(title, handle):
    q="""query($handle:String!){ collectionByIdentifier(identifier:{handle:$handle}){ id title } }"""
    existing=shopify_graphql(q,{"handle":handle}).get("collectionByIdentifier")
    if existing: return existing["id"], False
    m="""mutation($input:CollectionInput!){ collectionCreate(input:$input){ collection{ id } userErrors{ field message } } }"""
    d=shopify_graphql(m,{"input":{"title":title,"handle":handle}})["collectionCreate"]
    if d.get("userErrors"): raise RuntimeError(json.dumps(d["userErrors"]))
    return d["collection"]["id"], True

def shopify_product_id_by_handle(handle):
    q="""query($handle:String!){ productByIdentifier(identifier:{handle:$handle}){ id } }"""
    p=shopify_graphql(q,{"handle":handle}).get("productByIdentifier")
    return p["id"] if p else None

def shopify_add_product_to_collection(product_id, collection_id):
    m="""mutation($product:ProductUpdateInput!){ productUpdate(product:$product){ product{ id } userErrors{ field message } } }"""
    d=shopify_graphql(m,{"product":{"id":product_id,"collectionsToJoin":[collection_id]}})["productUpdate"]
    if d.get("userErrors"): raise RuntimeError(json.dumps(d["userErrors"]))

def shopify_sync_collections():
    products=load_products()
    results={"created_collections":[],"matched":0,"not_found_in_shopify":[],"errors":[]}
    collection_cache={}
    for p in products:
        if not p.get("approved") or p.get("skipped"): continue
        subcategory=clean(p.get("subcategory"))
        if not subcategory: continue
        handle=slugify(subcategory)
        if handle not in collection_cache:
            try:
                cid, created=shopify_find_or_create_collection(subcategory, handle)
                collection_cache[handle]=cid
                if created: results["created_collections"].append(subcategory)
            except Exception as e:
                results["errors"].append({"subcategory":subcategory,"error":str(e)})
                continue
        collection_id=collection_cache[handle]
        try:
            product_id=shopify_product_id_by_handle(p["handle"])
            if not product_id:
                results["not_found_in_shopify"].append(p["handle"]); continue
            shopify_add_product_to_collection(product_id, collection_id)
            results["matched"]+=1
        except Exception as e:
            results["errors"].append({"handle":p["handle"],"error":str(e)})
    return results

def taxonomy_map_path(): return os.path.join(DATA_DIR,"shopify_taxonomy_map.json")
def load_taxonomy_map():
    return json.load(open(taxonomy_map_path(),encoding="utf-8")) if os.path.exists(taxonomy_map_path()) else {"bead_shape":{},"color":{}}

COLOR_BASE_SWATCH_GIDS={
    "Beige":"gid://shopify/TaxonomyValue/6","Black":"gid://shopify/TaxonomyValue/1","Blue":"gid://shopify/TaxonomyValue/2",
    "Bronze":"gid://shopify/TaxonomyValue/657","Brown":"gid://shopify/TaxonomyValue/7","Clear":"gid://shopify/TaxonomyValue/17",
    "Gold":"gid://shopify/TaxonomyValue/4","Gray":"gid://shopify/TaxonomyValue/8","Green":"gid://shopify/TaxonomyValue/9",
    "Multicolor":"gid://shopify/TaxonomyValue/2865","Navy":"gid://shopify/TaxonomyValue/15","Orange":"gid://shopify/TaxonomyValue/10",
    "Pink":"gid://shopify/TaxonomyValue/11","Purple":"gid://shopify/TaxonomyValue/12","Red":"gid://shopify/TaxonomyValue/13",
    "Rose gold":"gid://shopify/TaxonomyValue/16","Silver":"gid://shopify/TaxonomyValue/5","White":"gid://shopify/TaxonomyValue/3",
    "Yellow":"gid://shopify/TaxonomyValue/14",
}
BEAD_SHAPE_BASE_GIDS={
    "Hair pipe":"gid://shopify/TaxonomyValue/14881","Heart":"gid://shopify/TaxonomyValue/17414","Other":"gid://shopify/TaxonomyValue/27273",
    "Oval":"gid://shopify/TaxonomyValue/17415","Round":"gid://shopify/TaxonomyValue/17416","Seed":"gid://shopify/TaxonomyValue/14882",
    "Square":"gid://shopify/TaxonomyValue/17417","Star":"gid://shopify/TaxonomyValue/14638","Tube":"gid://shopify/TaxonomyValue/17418",
}
SOLID_PATTERN_GID="gid://shopify/TaxonomyValue/2874"
COLOR_FINISH_WORDS={"TRANSPARENT","TRANSPAENT","TRANSPAR","TRANS","PARENT","OPAQUE","MATTE"}
CODE_PREFIX_RE=re.compile(r"^[0-9A-Za-z]{4,6}\s*-\s*")

def normalize_color_key(label):
    s=CODE_PREFIX_RE.sub("", label or "")
    s=s.replace("&amp;","&").replace("&nbsp;"," ")
    s=re.sub(r"[^A-Za-z0-9]+"," ", s).upper().strip()
    return " ".join(w for w in s.split() if w not in COLOR_FINISH_WORDS)

def shopify_list_metaobjects(mo_type):
    nodes, cursor = [], None
    while True:
        q="""query($type:String!,$after:String){ metaobjects(type:$type, first:100, after:$after){ edges{ cursor node{ id handle displayName } } pageInfo{ hasNextPage } } }"""
        data=shopify_graphql(q,{"type":mo_type,"after":cursor})
        edges=data["metaobjects"]["edges"]
        nodes.extend(e["node"] for e in edges)
        if not data["metaobjects"]["pageInfo"]["hasNextPage"] or not edges: break
        cursor=edges[-1]["cursor"]
    return nodes

def shopify_create_color_pattern_metaobject(label, base_color_name):
    base_gid=COLOR_BASE_SWATCH_GIDS.get(base_color_name)
    if not base_gid: raise RuntimeError(f"No base swatch GID for {base_color_name!r}")
    m="""mutation($obj:MetaobjectCreateInput!){ metaobjectCreate(metaobject:$obj){ metaobject{ handle } userErrors{ field message } } }"""
    d=shopify_graphql(m,{"obj":{"type":"shopify--color-pattern","fields":[
        {"key":"label","value":label},
        {"key":"color_taxonomy_reference","value":json.dumps([base_gid])},
        {"key":"pattern_taxonomy_reference","value":SOLID_PATTERN_GID},
    ]}})["metaobjectCreate"]
    if d.get("userErrors"): raise RuntimeError(json.dumps(d["userErrors"]))
    return d["metaobject"]["handle"]

def shopify_create_bead_shape_metaobject(label, base_shape_name):
    base_gid=BEAD_SHAPE_BASE_GIDS.get(base_shape_name)
    if not base_gid: raise RuntimeError(f"No base bead shape GID for {base_shape_name!r}")
    m="""mutation($obj:MetaobjectCreateInput!){ metaobjectCreate(metaobject:$obj){ metaobject{ handle } userErrors{ field message } } }"""
    d=shopify_graphql(m,{"obj":{"type":"shopify--bead-shape","fields":[
        {"key":"label","value":label},
        {"key":"taxonomy_reference","value":base_gid},
    ]}})["metaobjectCreate"]
    if d.get("userErrors"): raise RuntimeError(json.dumps(d["userErrors"]))
    return d["metaobject"]["handle"]

class MetaobjectResolver:
    """Resolves Margola color/bead-shape text to Shopify metaobject handles for CSV export, creating new library entries when no match exists. Caches lookups for the lifetime of one export."""
    def __init__(self):
        self.taxonomy_map=load_taxonomy_map()
        self._color_by_key=None
        self._shape_by_label=None
        self._created_colors={}
        self._created_shapes={}
        self.created_colors=[]
        self.created_shapes=[]
        self.unmapped_colors=[]

    def _colors(self):
        if self._color_by_key is None:
            self._color_by_key={}
            for n in shopify_list_metaobjects("shopify--color-pattern"):
                self._color_by_key.setdefault(normalize_color_key(n["displayName"]), []).append(n)
        return self._color_by_key

    def _shapes(self):
        if self._shape_by_label is None:
            self._shape_by_label={n["displayName"].strip().upper():n for n in shopify_list_metaobjects("shopify--bead-shape")}
        return self._shape_by_label

    def color_handle(self, color_name):
        color_name=clean(color_name)
        if not color_name: return ""
        key=normalize_color_key(color_name)
        by_key=self._colors()
        if key in by_key:
            candidates=by_key[key]
            plain=[c for c in candidates if not CODE_PREFIX_RE.match(c["displayName"])]
            return (plain[0] if plain else candidates[0])["handle"]
        if key in self._created_colors: return self._created_colors[key]
        base_color=self.taxonomy_map.get("color",{}).get(color_name)
        if not base_color:
            self.unmapped_colors.append(color_name); return ""
        label=key.title()
        handle=shopify_create_color_pattern_metaobject(label, base_color)
        self._created_colors[key]=handle
        self.created_colors.append(label)
        return handle

    def bead_shape_handle(self, bead_shape):
        spec=self.taxonomy_map.get("bead_shape",{}).get(clean(bead_shape))
        if not spec: return ""
        label=spec["label"]; key=label.strip().upper()
        by_label=self._shapes()
        if key in by_label: return by_label[key]["handle"]
        if key in self._created_shapes: return self._created_shapes[key]
        handle=shopify_create_bead_shape_metaobject(label, spec["base"])
        self._created_shapes[key]=handle
        self.created_shapes.append(label)
        return handle

class Handler(BaseHTTPRequestHandler):
    def send_json(self, obj, status=200):
        data = json.dumps(obj).encode("utf-8")
        self.send_response(status); self.send_header("Content-Type","application/json"); self.send_header("Content-Length",str(len(data))); self.end_headers(); self.wfile.write(data)
    def read_body(self): return self.rfile.read(int(self.headers.get("Content-Length","0")))
    def send_csv(self, text, filename):
        data = text.encode("utf-8-sig")
        self.send_response(200); self.send_header("Content-Type","text/csv; charset=utf-8"); self.send_header("Content-Disposition",f"attachment; filename={filename}"); self.send_header("Content-Length",str(len(data))); self.end_headers(); self.wfile.write(data)
    def serve_static(self,path):
        full = os.path.join(APP_DIR, path.lstrip("/"))
        if not os.path.abspath(full).startswith(APP_DIR) or not os.path.exists(full): self.send_response(404); self.end_headers(); return
        data = open(full,"rb").read()
        self.send_response(200); self.send_header("Content-Type",guess(path)); self.send_header("Content-Length",str(len(data))); self.end_headers(); self.wfile.write(data)
    def do_GET(self):
        path = urlparse(self.path).path
        if path == "/": return self.serve_static("/static/index.html")
        if path.startswith("/static/"): return self.serve_static(path)
        products = load_products()
        if path == "/api/products": return self.send_json({"summary":summary(products),"products":products})
        if path == "/api/export/review.csv": return self.send_csv(review_csv(products),"margola_review_report.csv")
        if path == "/api/export/shopify-preview.csv":
            try: return self.send_csv(shopify_csv(products, True),"margola_shopify_preview.csv")
            except Exception as e: return self.send_json({"ok":False,"error":str(e)},500)
        if path == "/api/export/shopify-test.csv":
            try: return self.send_csv(shopify_csv(products, True, 3),"margola_shopify_TEST_3_products.csv")
            except Exception as e: return self.send_json({"ok":False,"error":str(e)},500)
        if path == "/api/shopify/config":
            c=shopify_config(); return self.send_json({"store":c["store"],"has_client_id":bool(c["client_id"]),"has_client_secret":bool(c["client_secret"])})
        if path == "/api/shopify/test":
            try: return self.send_json({"ok":True,"shop":shopify_test_connection()})
            except Exception as e: return self.send_json({"ok":False,"error":str(e)},500)
        if path == "/api/shopify/files":
            try:
                files=shopify_list_files(); save_file_map(files); matched,total=apply_file_matches_to_products(); return self.send_json({"ok":True,"files":files,"matched":matched,"total_products":total})
            except Exception as e: return self.send_json({"ok":False,"error":str(e)},500)
        if path == "/api/photos/resolve-from-drive":
            index = load_drive_index()
            if not index["found_any"]:
                return self.send_json({"ok":False,"error":f"No drive filelist found. Expected one of: {DRIVE_FILELISTS}"},400)
            proposals=[]
            for p in products:
                if p.get("image_src"): continue
                found = resolve_photo_from_drive(p, index)
                if found:
                    proposals.append({"id":p["id"],"title":p.get("title"),"factory_style":p.get("factory_style"),
                                       "source_path":found["source_path"],"target_filename":found["target_filename"],
                                       "reused_other_size":found["reused_other_size"]})
            return self.send_json({"ok":True,"proposals":proposals,"filelists":DRIVE_FILELISTS})
        self.send_response(404); self.end_headers()
    def do_POST(self):
        path = urlparse(self.path).path
        if path == "/api/shopify/sync-collections":
            try: return self.send_json({"ok":True, **shopify_sync_collections()})
            except Exception as e: return self.send_json({"ok":False,"error":str(e)},500)
        if path == "/api/photos/apply-drive-matches":
            data = json.loads(self.read_body().decode("utf-8")); matches = data.get("matches",[])
            files = load_file_map(); existing_by_filename = {f.get("filename"):f for f in files if f.get("filename")}
            uploaded=[]; errors=[]
            for m in matches:
                source_path, target_filename = m.get("source_path",""), m.get("target_filename","")
                if not source_path or not target_filename:
                    errors.append({"target_filename":target_filename,"error":"Missing source_path or target_filename"}); continue
                if target_filename in existing_by_filename:
                    uploaded.append({"target_filename":target_filename,"status":"already_known"}); continue
                if not os.path.exists(source_path):
                    errors.append({"target_filename":target_filename,"error":f"Drive not connected or file moved: {source_path}"}); continue
                try:
                    result = shopify_upload_file(source_path, alt=target_filename, filename=target_filename)
                    files.append(result); uploaded.append({"target_filename":target_filename,"url":result.get("url")})
                except Exception as e:
                    errors.append({"target_filename":target_filename,"error":str(e)})
            save_file_map(files); matched,total = apply_file_matches_to_products()
            return self.send_json({"ok":len(errors)==0,"uploaded_count":len(uploaded),"error_count":len(errors),
                                    "uploaded":uploaded,"errors":errors,"matched":matched,"total_products":total})
        if path == "/api/reset":
            save_products([]); return self.send_json({"ok":True,"summary":summary([])})
        if path == "/api/product/update":
            data = json.loads(self.read_body().decode("utf-8")); products = load_products()
            for p in products:
                if p["id"] == data.get("id"): p.update(data.get("updates",{})); break
            save_products(products); return self.send_json({"ok":True,"summary":summary(products)})
        if path == "/api/descriptions/import":
            data = json.loads(self.read_body().decode("utf-8")); descs = parse_descriptions(data.get("text","")); products = load_products(); matched = 0
            for p in products:
                if p["handle"] in descs: p["generated_description"] = descs[p["handle"]]; p["description_approved"] = False; matched += 1
            save_products(products); return self.send_json({"ok":True,"matched":matched,"total_descriptions":len(descs),"summary":summary(products)})
        if path == "/api/shopify/upload-test-image":
            ctype=self.headers.get("Content-Type",""); m=re.search(r"boundary=(.+)", ctype)
            if not m: return self.send_json({"error":"Missing upload boundary"},400)

            body=self.read_body()
            boundary=("--"+m.group(1)).encode()
            uploaded=[]

            for part in body.split(boundary):
                if b'name="file"' not in part:
                    continue
                fm=re.search(rb'filename="([^"]+)"',part)
                if not fm:
                    continue
                filename=fm.group(1).decode("utf-8",errors="replace")
                if not filename:
                    continue
                pieces=part.split(b"\r\n\r\n",1)
                if len(pieces)!=2:
                    continue
                file_bytes=pieces[1].rstrip(b"\r\n--")
                if not file_bytes:
                    continue

                safe_name=os.path.basename(filename)
                upload_path=os.path.join(UPLOAD_DIR,"shopify_upload_"+safe_name)
                open(upload_path,"wb").write(file_bytes)
                uploaded.append((safe_name, upload_path))

            if not uploaded:
                return self.send_json({"error":"No images uploaded"},400)

            try:
                files=load_file_map()
                results=[]
                errors=[]

                existing_by_filename={f.get("filename"):f for f in files if f.get("filename")}

                for safe_name, upload_path in uploaded:
                    if safe_name in existing_by_filename:
                        results.append({"filename": safe_name, "status": "already_known", "url": existing_by_filename[safe_name].get("url","")})
                        continue

                    try:
                        result=shopify_upload_file(upload_path, alt=safe_name)
                        files.append(result)
                        results.append(result)
                    except Exception as item_error:
                        errors.append({"filename": safe_name, "error": str(item_error)})

                save_file_map(files)
                matched,total=apply_file_matches_to_products()

                return self.send_json({
                    "ok": len(errors)==0,
                    "uploaded_count": len(results),
                    "error_count": len(errors),
                    "files": results,
                    "errors": errors,
                    "matched": matched,
                    "total_products": total
                })
            except Exception as e:
                return self.send_json({"ok":False,"error":str(e)},500)
        if path == "/api/import":
            ctype = self.headers.get("Content-Type",""); m = re.search(r"boundary=(.+)", ctype)
            if not m: return self.send_json({"error":"Missing upload boundary"},400)
            body = self.read_body(); boundary = ("--"+m.group(1)).encode(); file_bytes = None; filename = "upload.xlsx"
            for part in body.split(boundary):
                if b'name="file"' in part:
                    fm = re.search(rb'filename="([^"]+)"', part)
                    if fm: filename = fm.group(1).decode("utf-8", errors="replace")
                    pieces = part.split(b"\r\n\r\n",1)
                    if len(pieces)==2: file_bytes = pieces[1].rstrip(b"\r\n--"); break
            if not file_bytes: return self.send_json({"error":"No file uploaded"},400)
            upload_path = os.path.join(UPLOAD_DIR, f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{slugify(filename)}.xlsx")
            open(upload_path,"wb").write(file_bytes)
            try:
                products = parse_xlsx(upload_path); save_products(products); return self.send_json({"summary":summary(products),"detected":detected_categories(products),"products":products})
            except Exception as e: return self.send_json({"error":str(e)},500)
        self.send_response(404); self.end_headers()

if __name__ == "__main__":
    port = int(os.environ.get("PORT","8787"))
    print(f"Margola Product Manager v3 running at http://localhost:{port}")
    ThreadingHTTPServer(("127.0.0.1", port), Handler).serve_forever()
