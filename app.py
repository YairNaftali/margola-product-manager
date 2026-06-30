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
    except Exception: return clean(v)

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

def infer(sheet_name, color_name):
    s = sheet_name.lower()
    c = clean(color_name).lower()
    if "crow" in s:
        return {"subcategory":"Crow Beads","bead_shape":"Crow Beads","color_type":"Opaque","factory_qty_standard":"1000 pieces","mini_qty_standard":"100 pieces"}
    if "roller" in s:
        if "transparent" in s or "transparent" in c or "transpaent" in c: ct = "Transparent"
        elif "opaque" in s or "opaque" in c: ct = "Opaque"
        else: ct = ""
        return {"subcategory":"Roller Beads","bead_shape":"Roller Beads","color_type":ct,"factory_qty_standard":"1200 beads","mini_qty_standard":"144 beads"}
    return {"subcategory":"","bead_shape":"","color_type":"","factory_qty_standard":"","mini_qty_standard":""}

def title_for(subcategory, size, color_name):
    # No assumptions: do not invent manufacturer/brand.
    return " ".join([x for x in ["Czech Glass", clean(subcategory), clean(size), clean(color_name)] if x])

def image_for(factory_style, color_name, subcategory, size):
    base = clean(factory_style) or " ".join([subcategory, size, color_name])
    return f"{slugify(base)}.jpg" if base else ""

def validate(p):
    checks = {
        "Title": bool(p.get("title")),
        "Handle": bool(p.get("handle")),
        "Mini price": bool(p.get("mini_price")),
        "Factory price": bool(p.get("factory_price")),
        "Mini weight": bool(p.get("mini_weight_oz")),
        "Factory weight": bool(p.get("factory_weight_oz")),
        "Mini SKU": bool(p.get("mini_style")),
        "Factory SKU": bool(p.get("factory_style")),
        "Collection": bool(p.get("collection")),
        "Subcategory": bool(p.get("subcategory")),
        "Color type": bool(p.get("color_type")),
        "Bead shape": bool(p.get("bead_shape")),
        "Image filename": bool(p.get("image_filename")),
        "Description": bool(p.get("generated_description") or p.get("source_description")),
    }
    score = int(round(sum(1 for ok in checks.values() if ok) / len(checks) * 100))
    warnings = [k for k, ok in checks.items() if not ok]
    if p.get("notes"): warnings.append(p["notes"])
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
            size = get_first(row, ["BEAD SIZE DIAMETER  MM","BEAD SIZE","SIZE"])
            color_number = get_first(row, ["COLOR NUMBER"])
            color_name = get_first(row, ["COLOR NAME"])
            source_description = get_first(row, ["DESCRIPTION"])
            factory_qty = get_first(row, ["FACTORY PACK UNIT QUANTITY"])
            factory_qty_desc = get_first(row, ["UNIT QUANTITY DESCRIPTION"])
            factory_price = get_first(row, ["UNIT PRICE PER FACTORY PACK","UNIT PRICE PER BAG"])
            factory_weight = get_first(row, ["WEIGHT PER FACTORY PACK"])
            mini_style = get_first(row, ["MINI PACK STYLE #","MINI PACK       STYLE #"])
            mini_qty = get_first(row, ["MINI PACK QUANTITY","MINI PACK QUANTITY DESCRIPTION"])
            mini_price = get_first(row, ["UNIT PRICE PER MINI PACK"])
            mini_weight = get_first(row, ["WEIGHT PER MINI PACK"])
            inf = infer(ws.title, color_name)
            title = title_for(inf["subcategory"], size, color_name)
            row_text = " ".join(clean(x) for x in vals).lower()
            notes = []
            if "missing phot" in row_text or "mising phot" in row_text: notes.append("Source note: missing photo")
            p = {
                "id":str(uuid.uuid4()), "source_file":os.path.basename(path), "source_sheet":ws.title, "source_row":row_num,
                "approved":False, "skipped":False, "status":"Needs Review",
                "title":title, "handle":slugify(title), "brand":"", "vendor":"Margola",
                "collection":"Czech Glass Beads", "subcategory":inf["subcategory"], "color_type":inf["color_type"], "bead_shape":inf["bead_shape"],
                "size":clean(size), "color_number":clean(color_number), "color_name":clean(color_name),
                "image_filename":image_for(factory_style,color_name,inf["subcategory"],size), "image_src":"", "image_alt":title,
                "factory_style":clean(factory_style), "factory_quantity":clean(factory_qty),
                "factory_quantity_description":clean(factory_qty_desc) or inf["factory_qty_standard"],
                "factory_price":money(factory_price), "factory_weight_oz":weight_oz(factory_weight),
                "mini_style":clean(mini_style), "mini_quantity":clean(mini_qty) or inf["mini_qty_standard"],
                "mini_price":money(mini_price), "mini_weight_oz":weight_oz(mini_weight),
                "source_description":clean(source_description), "generated_description":"", "description_approved":False,
                "notes":"; ".join(notes),
            }
            products.append(p)
    handles, skus = {}, {}
    for p in products:
        handles.setdefault(p["handle"], []).append(p["id"])
        for sku in [p["factory_style"], p["mini_style"]]:
            if sku: skus.setdefault(sku, []).append(p["id"])
    dup_handles = {k for k,v in handles.items() if len(v)>1}
    dup_skus = {k for k,v in skus.items() if len(v)>1}
    for p in products:
        more = []
        if p["handle"] in dup_handles: more.append("Duplicate handle")
        if p["factory_style"] in dup_skus: more.append("Duplicate factory SKU")
        if p["mini_style"] in dup_skus: more.append("Duplicate mini SKU")
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

def review_csv(products):
    fields = ["status","approved","skipped","validation_score","warnings","title","handle","collection","subcategory","color_type","bead_shape","size","color_number","color_name","image_filename","image_src","factory_style","factory_price","factory_weight_oz","mini_style","mini_price","mini_weight_oz","source_sheet","source_row"]
    out = io.StringIO(); w = csv.DictWriter(out, fieldnames=fields); w.writeheader()
    for p in products:
        v = p.get("validation",{})
        row = {f:p.get(f,"") for f in fields}; row["validation_score"] = v.get("score",""); row["warnings"] = "; ".join(v.get("warnings",[]))
        w.writerow(row)
    return out.getvalue()

def shopify_rows(products, approved_only=True, limit=None):
    rows, count = [], 0
    for p in products:
        if approved_only and (not p.get("approved") or p.get("skipped")): continue
        if limit is not None and count >= limit: break
        count += 1
        body = p.get("generated_description") or p.get("source_description") or ""
        common = {
            "Handle":p["handle"], "Vendor":p.get("vendor") or "Margola",
            "Product Category":"Arts & Entertainment > Hobbies & Creative Arts > Arts & Crafts > Art & Crafting Materials > Beads",
            "Type":p["subcategory"] or "Czech Glass Beads",
            "Tags":", ".join(x for x in [p["collection"],p["subcategory"],p["color_type"],p["bead_shape"],p["size"]] if x),
            "Published":"TRUE", "Option1 Name":"Bundle Pack Options", "Option1 Linked To":"product.metafields.custom.bundle_pack_options",
            "Variant Inventory Tracker":"shopify", "Variant Inventory Qty":"0", "Variant Inventory Policy":"deny", "Variant Fulfillment Service":"manual",
            "Variant Requires Shipping":"TRUE", "Variant Taxable":"TRUE", "Variant Weight Unit":"oz", "Status":"active",
            "Style Number (product.metafields.custom.style_number)":p["factory_style"],
            "Color Type (product.metafields.custom.color_type)":p["color_type"],
            "Bead shape (product.metafields.custom.bead_shape)":p["bead_shape"],
            "Color (product.metafields.custom.color)":p["color_name"],
            "Size (product.metafields.custom.size)":p["size"],
            "Price Description (product.metafields.custom.price_description)":p["factory_quantity_description"],
        }
        first = dict(common); first.update({"Title":p["title"],"Body (HTML)":body,"Option1 Value":"mini-pack","Variant SKU":p["mini_style"],"Variant Price":p["mini_price"],"Variant Grams":p["mini_weight_oz"],"Image Src":p.get("image_src",""),"Image Alt Text":p.get("image_alt") or p["title"],"Bundle Pack Options (product.metafields.custom.bundle_pack_options)":"mini-pack; factory-pack"})
        second = dict(common); second.update({"Title":"","Body (HTML)":"","Option1 Value":"factory-pack","Variant SKU":p["factory_style"],"Variant Price":p["factory_price"],"Variant Grams":p["factory_weight_oz"],"Image Src":"","Image Alt Text":"","Bundle Pack Options (product.metafields.custom.bundle_pack_options)":""})
        rows += [first, second]
    return rows

def shopify_csv(products, approved_only=True, limit=None):
    fields = ["Handle","Title","Body (HTML)","Vendor","Product Category","Type","Tags","Published","Option1 Name","Option1 Value","Option1 Linked To","Variant SKU","Variant Price","Variant Grams","Variant Weight Unit","Variant Inventory Tracker","Variant Inventory Qty","Variant Inventory Policy","Variant Fulfillment Service","Variant Requires Shipping","Variant Taxable","Image Src","Image Alt Text","Status","Bundle Pack Options (product.metafields.custom.bundle_pack_options)","Style Number (product.metafields.custom.style_number)","Color Type (product.metafields.custom.color_type)","Bead shape (product.metafields.custom.bead_shape)","Color (product.metafields.custom.color)","Size (product.metafields.custom.size)","Price Description (product.metafields.custom.price_description)"]
    out = io.StringIO(); w = csv.DictWriter(out, fieldnames=fields); w.writeheader()
    for row in shopify_rows(products, approved_only, limit): w.writerow(row)
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
    products=load_products(); files=load_file_map(); by={f.get("filename"):f.get("url") for f in files if f.get("filename") and f.get("url")}; matched=0
    for p in products:
        if p.get("image_filename") in by: p["image_src"]=by[p["image_filename"]]; matched+=1
    save_products(products); return matched,len(products)
def multipart_form_data(fields, files):
    boundary="----MargolaBoundary"+uuid.uuid4().hex; chunks=[]
    for k,v in fields.items(): chunks += [f"--{boundary}\r\n".encode(), f'Content-Disposition: form-data; name="{k}"\r\n\r\n'.encode(), str(v).encode(), b"\r\n"]
    for k,(filename,ctype,data) in files.items(): chunks += [f"--{boundary}\r\n".encode(), f'Content-Disposition: form-data; name="{k}"; filename="{filename}"\r\n'.encode(), f"Content-Type: {ctype}\r\n\r\n".encode(), data, b"\r\n"]
    chunks.append(f"--{boundary}--\r\n".encode()); return boundary,b"".join(chunks)
def shopify_upload_file(filepath, alt=""):
    filename=os.path.basename(filepath); ctype=mimetypes.guess_type(filepath)[0] or "image/jpeg"; size=str(os.path.getsize(filepath))
    q1="""mutation stagedUploadsCreate($input:[StagedUploadInput!]!){ stagedUploadsCreate(input:$input){ stagedTargets{ url resourceUrl parameters{ name value } } userErrors{ field message } } }"""
    d=shopify_graphql(q1,{"input":[{"filename":filename,"mimeType":ctype,"httpMethod":"POST","resource":"FILE","fileSize":size}]})["stagedUploadsCreate"]
    if d.get("userErrors"): raise RuntimeError(json.dumps(d["userErrors"]))
    t=d["stagedTargets"][0]; fields={p["name"]:p["value"] for p in t["parameters"]}; data=open(filepath,"rb").read(); boundary,body=multipart_form_data(fields,{"file":(filename,ctype,data)})
    req=urllib.request.Request(t["url"],data=body,headers={"Content-Type":f"multipart/form-data; boundary={boundary}"},method="POST")
    urllib.request.urlopen(req,timeout=120).read()
    q2="""mutation fileCreate($files:[FileCreateInput!]!){ fileCreate(files:$files){ files{ id alt fileStatus createdAt ... on MediaImage{ image{ url } } ... on GenericFile{ url } } userErrors{ field message } } }"""
    c=shopify_graphql(q2,{"files":[{"alt":alt or filename,"contentType":"IMAGE","originalSource":t["resourceUrl"]}]})["fileCreate"]
    if c.get("userErrors"): raise RuntimeError(json.dumps(c["userErrors"]))
    f=c["files"][0]; url=(f.get("image") or {}).get("url") or f.get("url") or ""
    return {"id":f.get("id",""),"filename":filename,"url":url,"alt":f.get("alt",""),"status":f.get("fileStatus",""),"createdAt":f.get("createdAt","")}

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
        if path == "/api/export/shopify-preview.csv": return self.send_csv(shopify_csv(products, True),"margola_shopify_preview.csv")
        if path == "/api/export/shopify-test.csv": return self.send_csv(shopify_csv(products, True, 3),"margola_shopify_TEST_3_products.csv")
        if path == "/api/shopify/config":
            c=shopify_config(); return self.send_json({"store":c["store"],"has_client_id":bool(c["client_id"]),"has_client_secret":bool(c["client_secret"])})
        if path == "/api/shopify/test":
            try: return self.send_json({"ok":True,"shop":shopify_test_connection()})
            except Exception as e: return self.send_json({"ok":False,"error":str(e)},500)
        if path == "/api/shopify/files":
            try:
                files=shopify_list_files(); save_file_map(files); matched,total=apply_file_matches_to_products(); return self.send_json({"ok":True,"files":files,"matched":matched,"total_products":total})
            except Exception as e: return self.send_json({"ok":False,"error":str(e)},500)
        self.send_response(404); self.end_headers()
    def do_POST(self):
        path = urlparse(self.path).path
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
            body=self.read_body(); boundary=("--"+m.group(1)).encode(); file_bytes=None; filename="upload.jpg"
            for part in body.split(boundary):
                if b'name="file"' in part:
                    fm=re.search(rb'filename="([^"]+)"',part)
                    if fm: filename=fm.group(1).decode("utf-8",errors="replace")
                    pieces=part.split(b"\r\n\r\n",1)
                    if len(pieces)==2: file_bytes=pieces[1].rstrip(b"\r\n--"); break
            if not file_bytes: return self.send_json({"error":"No image uploaded"},400)
            upload_path=os.path.join(UPLOAD_DIR,"shopify_test_"+os.path.basename(filename)); open(upload_path,"wb").write(file_bytes)
            try:
                result=shopify_upload_file(upload_path, alt=os.path.basename(filename)); files=load_file_map(); files.append(result); save_file_map(files); matched,total=apply_file_matches_to_products(); return self.send_json({"ok":True,"file":result,"matched":matched,"total_products":total})
            except Exception as e: return self.send_json({"ok":False,"error":str(e)},500)
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
                products = parse_xlsx(upload_path); save_products(products); return self.send_json({"summary":summary(products),"products":products})
            except Exception as e: return self.send_json({"error":str(e)},500)
        self.send_response(404); self.end_headers()

if __name__ == "__main__":
    port = int(os.environ.get("PORT","8787"))
    print(f"Margola Product Manager v3 running at http://localhost:{port}")
    ThreadingHTTPServer(("127.0.0.1", port), Handler).serve_forever()
