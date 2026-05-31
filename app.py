#!/usr/bin/env python3
"""
RICHGYAN Receipt Generator - Production Web App
Flask backend: DOCX + PDF output, Groq AI autofill
"""

import os
import json
import threading
import requests
from io import BytesIO
from datetime import datetime
from flask import Flask, request, jsonify, send_file, render_template

import tempfile
import subprocess
import platform

from docx import Document
from docx.shared import Pt, RGBColor

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

# ─── Config ────────────────────────────────────────────────────────────────────
BASE_DIR      = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, 'RICHGYAN_receipt.docx')
COUNTER_FILE  = os.path.join(BASE_DIR, 'receipt_counter.json')
RECEIPTS_DIR  = os.path.join(BASE_DIR, 'receipts')
GROQ_API_URL  = "https://api.groq.com/openai/v1/chat/completions"

os.makedirs(RECEIPTS_DIR, exist_ok=True)
counter_lock = threading.Lock()

# ─── Counter ───────────────────────────────────────────────────────────────────
def load_counter():
    if os.path.exists(COUNTER_FILE):
        try:
            with open(COUNTER_FILE, 'r') as f:
                return json.load(f).get('counter', 0)
        except Exception:
            return 0
    return 0

def save_counter(val):
    with open(COUNTER_FILE, 'w') as f:
        json.dump({'counter': val}, f)

def get_receipt_id():
    with counter_lock:
        c = load_counter() + 1
        save_counter(c)
        return f"RG-{datetime.now().year}-{str(c).zfill(4)}"

# ─── Shared helpers ────────────────────────────────────────────────────────────
def number_to_words(num):
    ones  = ['','One','Two','Three','Four','Five','Six','Seven','Eight','Nine']
    tens  = ['','','Twenty','Thirty','Forty','Fifty','Sixty','Seventy','Eighty','Ninety']
    teens = ['Ten','Eleven','Twelve','Thirteen','Fourteen','Fifteen',
             'Sixteen','Seventeen','Eighteen','Nineteen']

    if num == 0: return 'Zero'

    def cvt(n):
        r = ''
        if n >= 100: r += ones[n//100]+' Hundred '; n %= 100
        if n >= 20:  r += tens[n//10]+' '; r += (ones[n%10]+' ' if n%10 else '')
        elif n >= 10: r += teens[n-10]+' '
        elif n > 0:  r += ones[n]+' '
        return r.strip()

    parts = []
    h = num % 1000;  num //= 1000
    if h: parts.insert(0, cvt(h))
    th = num % 100;  num //= 100
    if th: parts.insert(0, cvt(th)+' Thousand')
    lk = num % 100;  num //= 100
    if lk: parts.insert(0, cvt(lk)+' Lakh')
    cr = num % 100
    if cr: parts.insert(0, cvt(cr)+' Crore')
    return ' '.join(parts).replace('  ',' ').strip()

def format_amount(amount):
    try: return f"\u20b9 {int(amount):,}"
    except: return str(amount)

def validate_fields(data):
    """Returns (name, amt_int, purpose, month, branch) or raises ValueError with dict."""
    errors = {}
    name    = data.get('name','').strip()
    purpose = data.get('purpose','').strip()
    month   = data.get('month','').strip()
    branch  = data.get('branch','').strip()
    if not name:    errors['name']    = 'Name is required'
    if not purpose: errors['purpose'] = 'Purpose is required'
    if not month:   errors['month']   = 'Month is required'
    if not branch:  errors['branch']  = 'Branch is required'
    try:
        amt = int(str(data.get('amount','')).replace(',','').replace('\u20b9','').strip())
        if amt <= 0: raise ValueError
    except Exception:
        errors['amount'] = 'Valid positive amount is required'
        amt = 0
    if errors:
        raise ValueError(errors)
    return name, amt, purpose, month, branch

# ═══════════════════════════════════════════════════════════════════════════════
# DOCX Generator
# ═══════════════════════════════════════════════════════════════════════════════
def create_receipt_docx(name, amount, purpose, month, branch, receipt_id, today):
    amount_fmt   = format_amount(amount)
    amount_words = number_to_words(amount)

    field_map = {
        'Received with thanks from': f'Received with thanks from {name}',
        'Amount \u2026':             f'Amount {amount_fmt}',
        'In Word \u2026':            f'In Word {amount_words}',
        'For \u2026':                f'For {purpose}',
        'Month \u2026':              f'Month {month}',
        'Branch \u2026':             f'Branch {branch}',
    }

    doc = Document(TEMPLATE_PATH)
    money_receipt_idx = None

    for idx, para in enumerate(doc.paragraphs):
        text = para.text

        if 'MONEY RECEIPT' in text:
            money_receipt_idx = idx
            continue

        if money_receipt_idx is not None and idx == money_receipt_idx + 1 and text.strip() == '':
            if para.runs:
                r = para.runs[0]
                r.text = f'Receipt No: {receipt_id}     Date: {today}'
                r.bold = True
                for extra in para.runs[1:]: extra.text = ''
            else:
                run = para.add_run(f'Receipt No: {receipt_id}     Date: {today}')
                run.bold = True
            money_receipt_idx = None
            continue

        for key, replacement in field_map.items():
            if text.startswith(key) or (key in text and '\u2026' in text):
                if para.runs:
                    first = para.runs[0]
                    saved_bold  = first.bold
                    saved_size  = first.font.size
                    saved_color = None
                    try: saved_color = first.font.color.rgb
                    except: pass
                    for r in para.runs: r.text = ''
                    first.text = replacement
                    first.bold = saved_bold if saved_bold is not None else True
                    if saved_size:  first.font.size = saved_size
                    if saved_color: first.font.color.rgb = saved_color
                else:
                    para.add_run(replacement)
                break

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# ═══════════════════════════════════════════════════════════════════════════════
# PDF Generator  (converts the DOCX via MS Word/LibreOffice — pixel-perfect match)
# ═══════════════════════════════════════════════════════════════════════════════
def create_receipt_pdf(name, amount, purpose, month, branch, receipt_id, today):
    """
    Generate a PDF by:
      1. Building the filled DOCX (same as /api/generate)
      2. Converting it to PDF with docx2pdf (uses MS Word on Windows,
         LibreOffice on Linux/Mac)
    This guarantees logo, signature, footer and ₹ symbol all match the Word doc.
    """
    # Step 1 – generate the filled DOCX into a temp file
    docx_buf = create_receipt_docx(name, amount, purpose, month, branch, receipt_id, today)

    tmp_dir = tempfile.mkdtemp()
    docx_path = os.path.join(tmp_dir, f'{receipt_id}.docx')
    pdf_path  = os.path.join(tmp_dir, f'{receipt_id}.pdf')

    try:
        with open(docx_path, 'wb') as f:
            f.write(docx_buf.read())

        # Step 2 – convert to PDF
        _convert_docx_to_pdf(docx_path, pdf_path)

        with open(pdf_path, 'rb') as f:
            buf = BytesIO(f.read())
        buf.seek(0)
        return buf

    finally:
        # Clean up temp files
        for p in (docx_path, pdf_path):
            try:
                if os.path.exists(p):
                    os.remove(p)
            except Exception:
                pass
        try:
            os.rmdir(tmp_dir)
        except Exception:
            pass


def _convert_docx_to_pdf(docx_path, pdf_path):
    """Convert a .docx file to .pdf using the best available tool."""

    # ── Option A: docx2pdf (MS Word on Windows, LibreOffice on Mac/Linux) ──
    try:
        from docx2pdf import convert
        convert(docx_path, pdf_path)
        if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
            return
    except Exception:
        pass

    # ── Option B: LibreOffice headless (Linux/Mac servers) ──
    libreoffice_bins = [
        'libreoffice', 'soffice',
        '/usr/bin/libreoffice', '/usr/bin/soffice',
        '/Applications/LibreOffice.app/Contents/MacOS/soffice',
    ]
    out_dir = os.path.dirname(pdf_path)
    for lo in libreoffice_bins:
        try:
            result = subprocess.run(
                [lo, '--headless', '--convert-to', 'pdf',
                 '--outdir', out_dir, docx_path],
                capture_output=True, timeout=60
            )
            expected = os.path.join(out_dir,
                os.path.splitext(os.path.basename(docx_path))[0] + '.pdf')
            if os.path.exists(expected) and os.path.getsize(expected) > 0:
                if expected != pdf_path:
                    os.rename(expected, pdf_path)
                return
        except (FileNotFoundError, subprocess.TimeoutExpired):
            continue

    raise RuntimeError(
        'PDF conversion failed. On Windows, make sure Microsoft Word is installed. '
        'On Linux/Mac servers, install LibreOffice: apt install libreoffice'
    )

# ═══════════════════════════════════════════════════════════════════════════════
# Groq AI
# ═══════════════════════════════════════════════════════════════════════════════
def call_groq(api_key, prompt_text):
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    body = {
        "model": "llama3-8b-8192",
        "messages": [
            {"role": "system", "content": (
                "You are a receipt assistant for RICHGYAN INDIA, a computer & vocational institute. "
                "Given a user description, extract receipt fields as JSON. "
                "Respond ONLY with a raw JSON object, no markdown, no explanation. "
                "Fields: name (string), amount (integer in rupees, no symbol), "
                "purpose (string, e.g. 'Basic Computer Course Fee'), "
                "month (string, e.g. 'June 2026'), branch (string, e.g. 'Guwahati'). "
                "If a field cannot be determined, use an empty string or 0."
            )},
            {"role": "user", "content": prompt_text}
        ],
        "temperature": 0.2, "max_tokens": 300,
    }
    resp = requests.post(GROQ_API_URL, headers=headers, json=body, timeout=15)
    resp.raise_for_status()
    content = resp.json()["choices"][0]["message"]["content"].strip()
    if content.startswith("```"):
        content = content.split("```")[1]
        if content.startswith("json"): content = content[4:]
    return json.loads(content)

# ═══════════════════════════════════════════════════════════════════════════════
# Routes
# ═══════════════════════════════════════════════════════════════════════════════
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/groq-autofill', methods=['POST'])
def groq_autofill():
    data    = request.get_json(force=True)
    api_key = data.get('api_key','').strip()
    prompt  = data.get('prompt','').strip()
    if not api_key: return jsonify({'error': 'Groq API key is required'}), 400
    if not prompt:  return jsonify({'error': 'Prompt text is required'}), 400
    try:
        fields = call_groq(api_key, prompt)
        return jsonify({'success': True, 'fields': fields})
    except requests.exceptions.HTTPError as e:
        return jsonify({'error': f'Groq API error: {e.response.status_code} – check your API key'}), 400
    except json.JSONDecodeError:
        return jsonify({'error': 'Groq returned unexpected format'}), 500
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/generate', methods=['POST'])
def generate():
    """Generate DOCX receipt."""
    data = request.get_json(force=True)
    try:
        name, amt, purpose, month, branch = validate_fields(data)
    except ValueError as e:
        return jsonify({'error': 'Validation failed', 'fields': e.args[0]}), 400
    try:
        receipt_id = get_receipt_id()
        today      = datetime.now().strftime('%d %B %Y')
        buf        = create_receipt_docx(name, amt, purpose, month, branch, receipt_id, today)
        return send_file(buf, as_attachment=True,
            download_name=f"{receipt_id}.docx",
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except Exception as e:
        return jsonify({'error': f'DOCX generation failed: {str(e)}'}), 500

@app.route('/api/generate-pdf', methods=['POST'])
def generate_pdf():
    """Generate PDF receipt."""
    data = request.get_json(force=True)
    try:
        name, amt, purpose, month, branch = validate_fields(data)
    except ValueError as e:
        return jsonify({'error': 'Validation failed', 'fields': e.args[0]}), 400
    try:
        receipt_id = get_receipt_id()
        today      = datetime.now().strftime('%d %B %Y')
        buf        = create_receipt_pdf(name, amt, purpose, month, branch, receipt_id, today)
        return send_file(buf, as_attachment=True,
            download_name=f"{receipt_id}.pdf",
            mimetype='application/pdf')
    except Exception as e:
        return jsonify({'error': f'PDF generation failed: {str(e)}'}), 500

@app.route('/api/generate-both', methods=['POST'])
def generate_both():
    """Generate both DOCX and PDF, return JSON with base64 encoded files."""
    import base64
    data = request.get_json(force=True)
    try:
        name, amt, purpose, month, branch = validate_fields(data)
    except ValueError as e:
        return jsonify({'error': 'Validation failed', 'fields': e.args[0]}), 400
    try:
        receipt_id = get_receipt_id()
        today      = datetime.now().strftime('%d %B %Y')
        docx_buf   = create_receipt_docx(name, amt, purpose, month, branch, receipt_id, today)
        pdf_buf    = create_receipt_pdf(name, amt, purpose, month, branch, receipt_id, today)
        return jsonify({
            'receipt_id': receipt_id,
            'date':       today,
            'docx': {
                'filename': f'{receipt_id}.docx',
                'data': base64.b64encode(docx_buf.read()).decode(),
            },
            'pdf': {
                'filename': f'{receipt_id}.pdf',
                'data': base64.b64encode(pdf_buf.read()).decode(),
            }
        })
    except Exception as e:
        return jsonify({'error': f'Generation failed: {str(e)}'}), 500

@app.route('/api/preview', methods=['POST'])
def preview():
    data    = request.get_json(force=True)
    name    = data.get('name','').strip() or '\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026'
    purpose = data.get('purpose','').strip() or '\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026'
    month   = data.get('month','').strip() or '\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026'
    branch  = data.get('branch','').strip() or '\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026'
    today   = datetime.now().strftime('%d %B %Y')
    with counter_lock:
        next_num = load_counter() + 1
    try:
        amt          = int(str(data.get('amount',0)).replace(',','').replace('\u20b9','').strip() or 0)
        amount_fmt   = format_amount(amt) if amt else '\u20b9 \u2026\u2026\u2026\u2026\u2026\u2026'
        amount_words = number_to_words(amt) if amt else '\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026'
    except Exception:
        amount_fmt   = '\u20b9 \u2026\u2026\u2026\u2026\u2026\u2026'
        amount_words = '\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026'
    return jsonify({
        'receipt_id':   f'RG-{datetime.now().year}-{str(next_num).zfill(4)}',
        'date':         today,
        'name':         name,
        'amount_fmt':   amount_fmt,
        'amount_words': amount_words,
        'purpose':      purpose,
        'month':        month,
        'branch':       branch,
    })

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
