import os
import io
import zipfile
import re
from datetime import datetime, timedelta

from numpy import size
import pandas as pd
from flask import Flask, render_template, request, jsonify, send_file, abort, redirect, url_for, session
from werkzeug.utils import secure_filename

import config
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader

from PIL import Image, ImageChops
from num2words import num2words


app = Flask(__name__)
app.secret_key = "CHANGE_ME_TO_SOMETHING_RANDOM"  # change later

# session lifetime 24 hours
app.permanent_session_lifetime = timedelta(hours=24)

# Ensure folders exist (and fix if accidentally created as files)
for path in [config.UPLOAD_DIR, config.UPLOAD_EXCEL_DIR, config.OUT_INVOICE_DIR, config.OUT_ZIP_DIR]:
    if os.path.exists(path) and not os.path.isdir(path):
        os.remove(path)
    os.makedirs(path, exist_ok=True)

# ---- Column aliases (your Excel headers -> canonical keys) ----
CANONICAL_COLS = {
    "srno": ["Sr.No", "Sr.No ", "Sr No", "Sr. No.", "Sr.No."],
    "expert_name": ["Expert's name", "Experts name", "Expert name", "Expert Name"],
    "phone": ["Phone No.", "Phone No", "Phone", "Mobile"],
    "email": ["Email Address", "Email", "Email Id", "E-mail"],
    "address": ["Address", "Address ", "Full Address"],
    "pan": ["Pancard", "PAN", "Pan Card"],
    "bank_details": ["Bank Details", "Bank", "Bank Name"],
    "account_no": ["Ac/No.", "Ac/No", "Account No", "Account Number", "A/c No"],
    "ifsc": ["IFSC Code", "IFSC", "IFSC code"],
    "total_sales": ["Total Sales", "Sales", "TotalSale"],
    "commission": ["Commission", "Comm"],
    "in_words": ["In words", "In Words", "Amount in words"],
    "commission_pct": ["% of Commission", "Commission %", "Percent of Commission"],
    "invoice_number": ["Invoice number", "Invoice No", "Invoice No."],
    "notes": ["Notes", "Remark", "Remarks"],
    "payment_status": ["Payment Status", "Status"],
}

MONTHS = ["January","February","March","April","May","June","July","August","September","October","November","December"]
YEARS = list(range(2024, 2031))

# ---------------- AUTH ----------------

LOGIN_ID = "sa"
LOGIN_PW = "sa123"

def login_required():
    return session.get("logged_in") is True

def require_login():
    if not login_required():
        return redirect(url_for("login"))
    return None

# ---------------- Excel helpers ----------------

def pick_col(df, options):
    for c in options:
        if c in df.columns:
            return c
    return None

def standardize_df(df):
    out = {}
    missing = []
    for canon, aliases in CANONICAL_COLS.items():
        actual = pick_col(df, aliases)
        if not actual:
            missing.append((canon, aliases))
        else:
            out[canon] = df[actual]

    if missing:
        msg = "Missing required columns:\n"
        for canon, aliases in missing:
            msg += f"- {canon} (any of {aliases})\n"
        msg += f"\nFound columns: {list(df.columns)}"
        raise ValueError(msg)

    return pd.DataFrame(out).fillna("")

def load_data():
    excel_path = session.get("excel_path")
    if not excel_path or not os.path.exists(excel_path):
        return None

    df = pd.read_excel(excel_path, sheet_name=config.SHEET_NAME, engine="openpyxl")
    sdf = standardize_df(df)
    for col in sdf.columns:
        sdf[col] = sdf[col].apply(clean_text)

    for col in ["total_sales", "commission", "account_no"]:
        sdf[col] = sdf[col].apply(lambda x: "" if str(x).strip().lower() in ("nan", "none") else x)

    return sdf

# ---------------- misc utils ----------------



def clean_text(s):
    if s is None:
        return ""
    s = str(s)

    # Replace NBSP with normal space
    s = s.replace("\u00A0", " ")

    # Remove zero-width / direction / formatting marks
    s = re.sub(r"[\u200B-\u200F\u202A-\u202E\u2060-\u206F]", "", s)

    # Remove common black-square glyphs if they exist
    s = s.replace("\u25A0", "").replace("\u25AA", "").replace("■", "")

    # Normalize whitespace
    s = re.sub(r"\s+", " ", s).strip()
    return s

def safe_filename(s: str) -> str:
    return secure_filename(str(s)).strip("_")

def signature_path_for(expert_name: str):
    base = safe_filename(expert_name)
    p = os.path.join(config.UPLOAD_DIR, base + ".png")
    if os.path.exists(p):
        return p
    for ext in (".jpg", ".jpeg"):
        p = os.path.join(config.UPLOAD_DIR, base + ext)
        if os.path.exists(p):
            return p
    return None

def to_money_number(x):
    if x is None:
        return None
    if isinstance(x, (int, float)) and pd.isna(x):
        return None
    s = str(x).strip()
    if s == "" or s.lower() == "nan":
        return None
    s = s.replace(",", "")
    try:
        return float(s)
    except:
        return None

def amount_in_words(row: dict) -> str:
    raw = str(row.get("in_words", "")).strip()
    if raw and raw.lower() not in ["nan", "none", "0", "0.0"]:
        return raw

    amt = to_money_number(row.get("commission", None))
    if amt is None:
        return ""

    amt_int = int(round(amt))
    words = num2words(amt_int, lang="en_IN").replace("-", " ")
    return words[:1].upper() + words[1:]

def fmt_money(x):
    n = to_money_number(x)
    if n is None:
        return str(x) if x is not None else ""
    if abs(n - round(n)) < 1e-9:
        return f"{int(round(n)):,}"
    return f"{n:,.2f}"

def fmt_account_no(x) -> str:
    s = str(x).strip()
    if s.lower() in ("nan", "none"):
        return ""
    if s.endswith(".0"):
        s = s[:-2]
    return s

def draw_signature(c, signature_path, x, y, box_w, box_h):
    """Draw signature cropped + aspect fit into box."""
    img = Image.open(signature_path)

    # Remove transparency to white bg
    if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
        bg = Image.new("RGB", img.size, (255, 255, 255))
        bg.paste(img.convert("RGBA"), mask=img.convert("RGBA").split()[-1])
        img = bg
    else:
        img = img.convert("RGB")

    # Crop white margins
    bg = Image.new("RGB", img.size, (255, 255, 255))
    diff = ImageChops.difference(img, bg)
    bbox = diff.getbbox()
    if bbox:
        img = img.crop(bbox)

    reader = ImageReader(img)

    iw, ih = img.size
    scale = min(box_w / iw, box_h / ih)
    draw_w = iw * scale
    draw_h = ih * scale

    dx = x + (box_w - draw_w) / 2
    dy = y + (box_h - draw_h) / 2
    c.drawImage(reader, dx, dy, width=draw_w, height=draw_h)

# ---------------- PDF ----------------

def draw_invoice_pdf(row: dict, signature_path: str | None = None) -> bytes:
    """
    Simple, classy, customer-issued TAX INVOICE
    Issuer = Expert (customer)
    Bill To = Nutritionalab Private Limited (your company)
    Date removed. Uses Month/Year (Period).
    """
    month = session.get("period_month", "February")
    year = session.get("period_year", "2026")

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4

    def font(size=10, bold=False):
        c.setFont("Helvetica-Bold" if bold else "Helvetica", size)

    def txt(x, y, s, size=10, bold=False):
        font(size, bold)
        c.drawString(x, y, clean_text(s))

    def rtxt(x, y, s, size=10, bold=False):
        font(size, bold)
        c.drawRightString(x, y, clean_text(s))


    def line(x1, y1, x2, y2):
        c.setLineWidth(0.8)
        c.line(x1, y1, x2, y2)

    left = 18 * mm
    right = w - 18 * mm
    top = h - 18 * mm

    # --- Header blocks ---
    y = top

    # Right meta (Invoice no + Period)
    rtxt(right, y - 6, f"Invoice No: {row.get('invoice_number','')}", size=10, bold=True)
    rtxt(right, y - 20, f"Period: {month} {year}", size=10)

    # Left Issuer (Expert)
    txt(left, y, str(row.get("expert_name", "")).upper(), size=14, bold=True); y -= 14
    txt(left, y, f"Address: {row.get('address','')}", size=9); y -= 11
    txt(left, y, f"Phone: {row.get('phone','')}", size=9); y -= 11
    txt(left, y, f"Email: {row.get('email','')}", size=9); y -= 11
    txt(left, y, f"PAN: {row.get('pan','')}", size=9)

    # Center Title
    font(14, True)
    c.drawCentredString(w/2, top - 52, "TAX INVOICE")

    line(left, top - 62, right, top - 62)

    # --- Bill To (Company) ---
    y = top - 82
    txt(left, y, "Bill To:", size=10, bold=True); y -= 14
    txt(left, y, config.COMPANY_NAME, size=11, bold=True); y -= 12
    for ln in config.COMPANY_ADDR_LINES:
        txt(left, y, ln, size=9); y -= 11

    y -= 10
    line(left, y, right, y)
    y -= 16

    # table header
    txt(left, y, "Sr.", size=9, bold=True)
    txt(left + 18*mm, y, "Description", size=9, bold=True)
    rtxt(right - 55*mm, y, "Total Sales", size=9, bold=True)
    rtxt(right, y, "Amount", size=9, bold=True)

    y -= 10
    line(left, y, right, y)

    # item row
    # item row
    y -= 18
    desc = f"Affiliate marketing - {month} {year}"   # ✅ ONLY month-year
    total_sales = fmt_money(row.get("total_sales", ""))
    commission = fmt_money(row.get("commission", ""))

    txt(left, y, "1", size=9)
    txt(left + 18*mm, y, desc, size=9)
    rtxt(right - 55*mm, y, total_sales, size=9)
    rtxt(right, y, commission, size=9)


    y -= 16
    line(left, y, right, y)

    # totals
    y -= 18
    rtxt(right - 55*mm, y, "Total", size=10, bold=True)
    rtxt(right, y, commission, size=10, bold=True)

    # words
    y -= 20
    words = amount_in_words(row)
    txt(left, y, f"Rupees: {words} only." if words else "Rupees:", size=9)

    # bank details (issuer)
    # bank details (issuer)
    y -= 28
    txt(left, y, "Bank Details:", size=10, bold=True); y -= 14   # ✅ removed (Issuer) + extra space
    txt(left, y, f"{row.get('bank_details','')}", size=9); y -= 11
    txt(left, y, f"Account No: {fmt_account_no(row.get('account_no',''))}", size=9); y -= 11
    txt(left, y, f"IFSC Code: {row.get('ifsc','')}", size=9)

    # --- Signature placed closer (use remaining space better) ---
    sig_box_w = 70 * mm
    sig_box_h = 22 * mm
    sig_x = right - sig_box_w

    # put signature just below bank details area, not at extreme bottom
    sig_y = max(35 * mm, y - 75)  # keeps it on page nicely

    if signature_path and os.path.exists(signature_path):
        try:
            draw_signature(c, signature_path, sig_x, sig_y + 10, sig_box_w, sig_box_h)
        except Exception as e:
            print("Signature draw failed:", e, "path=", signature_path)

    # ✅ Center-align label under signature image
    font(9, False)
    c.drawCentredString(sig_x + (sig_box_w / 2), sig_y, "Authorised Signatory")


    c.showPage()
    c.save()
    buf.seek(0)
    return buf.read()

# ---------------- ROUTES ----------------

@app.get("/login")
def login():
    return render_template("login.html")

@app.post("/login")
def login_post():
    login_id = request.form.get("login_id","").strip()
    password = request.form.get("password","").strip()

    if login_id == LOGIN_ID and password == LOGIN_PW:
        session.permanent = True
        session["logged_in"] = True
        # default period
        session.setdefault("period_month", "February")
        session.setdefault("period_year", "2026")
        return redirect(url_for("upload_master"))
    return render_template("login.html", error="Invalid credentials")

@app.get("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.get("/upload")
def upload_master():
    go = require_login()
    if go: return go
    return render_template("upload.html")

@app.post("/upload")
def upload_master_post():
    go = require_login()
    if go: return go

    if "file" not in request.files:
        return render_template("upload.html", error="Please select an Excel file.")

    f = request.files["file"]
    if f.filename == "":
        return render_template("upload.html", error="Please select an Excel file.")

    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in [".xlsx", ".xlsm"]:
        return render_template("upload.html", error="Only .xlsx or .xlsm allowed.")

    fname = secure_filename(f.filename)
    save_path = os.path.join(config.UPLOAD_EXCEL_DIR, fname)
    f.save(save_path)

    session["excel_path"] = save_path
    return redirect(url_for("index"))

@app.post("/api/set-period")
def api_set_period():
    go = require_login()
    if go: return jsonify({"error":"not_logged_in"}), 401

    month = request.form.get("month","").strip()
    year = request.form.get("year","").strip()

    if month not in MONTHS:
        return jsonify({"error":"Invalid month"}), 400
    if year.isdigit() and int(year) in YEARS:
        session["period_month"] = month
        session["period_year"] = str(int(year))
        return jsonify({"ok": True})
    return jsonify({"error":"Invalid year"}), 400

@app.get("/")
def index():
    go = require_login()
    if go: return go

    df = load_data()
    if df is None:
        return redirect(url_for("upload_master"))

    cards = []
    for _, r in df.iterrows():
        row = r.to_dict()
        name = str(row.get("expert_name", "")).strip()
        row["has_signature"] = signature_path_for(name) is not None
        cards.append(row)

    return render_template(
        "index.html",
        experts=cards,
        months=MONTHS,
        years=YEARS,
        sel_month=session.get("period_month","February"),
        sel_year=session.get("period_year","2026")
    )

@app.post("/api/upload-signature")
def upload_signature():
    go = require_login()
    if go: return jsonify({"error":"not_logged_in"}), 401

    name = request.form.get("name", "").strip()
    if not name:
        return jsonify({"error": "Missing name"}), 400

    if "file" not in request.files:
        return jsonify({"error": "Missing file"}), 400

    f = request.files["file"]
    if f.filename == "":
        return jsonify({"error": "Empty filename"}), 400

    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in [".png", ".jpg", ".jpeg", ".webp"]:
        return jsonify({"error": "Only png/jpg/jpeg/webp allowed"}), 400

    out_name = safe_filename(name) + ".png"
    out_path = os.path.join(config.UPLOAD_DIR, secure_filename(out_name))

    try:
        img = Image.open(f.stream)

        # convert to RGB and remove transparency
        if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
            bg = Image.new("RGB", img.size, (255, 255, 255))
            bg.paste(img.convert("RGBA"), mask=img.convert("RGBA").split()[-1])
            img = bg
        else:
            img = img.convert("RGB")

        img.save(out_path, format="PNG")
    except Exception as e:
        return jsonify({"error": f"Could not process image: {e}"}), 400

    return jsonify({"ok": True, "path": out_name})

@app.get("/invoice.pdf")
def invoice_pdf():
    go = require_login()
    if go: return go

    name = request.args.get("name", "").strip()
    if not name:
        abort(400, "Missing name")

    df = load_data()
    if df is None:
        return redirect(url_for("upload_master"))

    matches = df[df["expert_name"].astype(str) == name]
    if matches.empty:
        abort(404, "Expert not found")

    row = matches.iloc[0].to_dict()
    sig = signature_path_for(name)
    pdf_bytes = draw_invoice_pdf(row, sig)

    inv_no = str(row.get("invoice_number", "")).strip()
    filename = safe_filename(f"{name}_Invoice_{inv_no}.pdf")

    return send_file(
        io.BytesIO(pdf_bytes),
        mimetype="application/pdf",
        as_attachment=True,
        download_name=filename,
    )

@app.get("/download-all.zip")
def download_all():
    go = require_login()
    if go: return go

    df = load_data()
    if df is None:
        return redirect(url_for("upload_master"))

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for _, r in df.iterrows():
            row = r.to_dict()
            name = str(row.get("expert_name", "")).strip()
            if not name:
                continue

            sig = signature_path_for(name)
            pdf_bytes = draw_invoice_pdf(row, sig)

            inv_no = str(row.get("invoice_number", "")).strip()
            pdf_name = safe_filename(f"{name}_Invoice_{inv_no}.pdf")
            zf.writestr(pdf_name, pdf_bytes)

    zip_buf.seek(0)
    return send_file(
        zip_buf,
        mimetype="application/zip",
        as_attachment=True,
        download_name="All_Invoices.zip",
    )

if __name__ == "__main__":
    app.run(debug=True)
