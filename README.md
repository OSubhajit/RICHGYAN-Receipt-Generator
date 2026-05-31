# RICHGYAN Receipt Generator — Production Web App

A production-ready web application that generates RICHGYAN INDIA money receipts as `.docx` files, with optional Groq AI autofill. Hosted once, usable by anyone on any device.

---

## ✨ Features

- **Pixel-perfect DOCX output** — same template, same formatting, same signs/symbols as original
- **Groq AI autofill** — describe a receipt in plain language, AI fills the form
- **Live preview** — see the receipt before downloading
- **Auto receipt numbering** — RG-2026-0001, RG-2026-0002, …
- **Amount in words** — automatic Indian-format (Lakh, Crore)
- **Any device, any browser** — fully responsive

---

## 📁 Project Structure

```
richgyan_web/
├── app.py                  ← Flask backend
├── RICHGYAN_receipt.docx   ← Original template (NEVER edit this)
├── requirements.txt
├── Procfile                ← For Render/Heroku
├── Dockerfile              ← For Docker/Railway
├── receipt_counter.json    ← Auto-created, tracks receipt numbers
├── receipts/               ← Temp folder (auto-created)
└── templates/
    └── index.html          ← Web UI
```

---

## 🚀 Hosting Options (Pick One)

### Option 1 — Render.com (FREE, Recommended)

1. Create a free account at https://render.com
2. Click **New → Web Service**
3. Connect your GitHub repo (upload this folder first)
4. Settings:
   - **Runtime**: Python 3
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app --bind 0.0.0.0:$PORT`
5. Click **Deploy** — done! You get a public URL like `https://richgyan-receipt.onrender.com`

### Option 2 — Railway.app (FREE tier)

1. Go to https://railway.app
2. New Project → Deploy from GitHub
3. It auto-detects Python and uses the `Procfile`
4. Done!

### Option 3 — Run Locally (LAN access)

```bash
# Install Python 3.10+, then:
pip install -r requirements.txt
python app.py
```

Open `http://YOUR_IP:5000` from any device on the same Wi-Fi.

### Option 4 — Docker

```bash
docker build -t richgyan-receipt .
docker run -p 5000:5000 richgyan-receipt
```

---

## 🤖 Groq AI Autofill Setup

1. Go to https://console.groq.com → Create a free account
2. API Keys → Create new key → Copy it
3. On the web app, paste the key in the **Groq API Key** field
4. Type your description → click **Autofill with AI**

**The key is never stored on the server** — it lives only in your browser session.

---

## 📄 Output Format

Generated DOCX files are identical in layout to the original template:

```
RICHGYAN INDIA
COMPUTER & VOCATIONAL INSTITUTE
(An I.T & Vocational Training Awareness Programme)
Registered under Ministry Of Corporate Affairs. GOVT. OF INDIA

         MONEY RECEIPT
Receipt No: RG-2026-0001   Date: 01 June 2026

Received with thanks from Subhajit Sarkar
Amount ₹ 2,000
In Word Two Thousand
For Basic Computer Course Fee
Month June 2026
Branch Guwahati

Thank you for your payment!
...
Authorized Signature
```

---

## ⚙️ Environment Variables (optional)

| Variable | Default | Purpose |
|----------|---------|---------|
| `PORT`   | 5000    | Port to run on |

---

## 🔒 Notes

- The `RICHGYAN_receipt.docx` template is **never modified** — a fresh copy is used per receipt
- Receipt counter persists in `receipt_counter.json` — back it up if you redeploy
- On Render free tier, the counter resets on redeploy (use a persistent disk add-on for production)
