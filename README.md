# Margola Product Manager v4

Adds Shopify client-credentials support.

Create `.env` inside this folder:

```env
SHOPIFY_STORE=ksbb0i-kh.myshopify.com
SHOPIFY_CLIENT_ID=your_client_id
SHOPIFY_CLIENT_SECRET=your_client_secret
```

Run:

```bash
cd ~/Downloads/Margola_Product_Manager_v4
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

Open `http://localhost:8787`.

Use the Shopify tab to check config, test connection, sync Shopify Files, or upload one test image.
