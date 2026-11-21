"""KeyStone Tutors - Streamlit site.

Run it with:   streamlit run app.py   (or double-click "Start KeyStone Site")

Everything the site saves (bookings, tutor applications, packages, reviews,
tutor profiles) goes to your Google Sheet when SHEET_URL and
[gcp_service_account] are set in the app's secrets. Without them (e.g. on
your laptop) it falls back to CSV files in the "data" folder.
Tutor transcripts go to Supabase Storage when SUPABASE_URL and SUPABASE_KEY
are set, otherwise to data/uploads.
"""
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

import pandas as pd
import streamlit as st

# ---------- Business details (edit these in one place) ----------
WHATSAPP = "27644067189"          # 064 406 7189 in international format
PHONE_DISPLAY = "064 406 7189"
EMAIL = "KeyStoneTutors@outlook.com"
INSTAGRAM = "keystone.tutors"
TUTOR_SHARE = 0.51                # tutors get 51% of the package value

PRICES = {
    # level: (single, 4-pack, 10-pack) per session
    "Primary (Gr 1-7)": (250, 200, 190),
    "Grade 8-9": (300, 250, 240),
    "Grade 10-11": (400, 350, 330),
    "Matric (Gr 12)": (450, 400, 375),
    "First-year uni": (500, 450, 420),
}
LEVELS = list(PRICES)
PACK_SIZES = {"Single lesson": 1, "4-lesson pack": 4, "10-lesson pack": 10}
# Subjects, grouped for the Home page. Add or remove freely.
SUBJECT_GROUPS = {
    "Maths & Sciences": ["Mathematics", "Mathematical Literacy", "Physical Sciences", "Life Sciences"],
    "Commerce": ["Accounting", "Economics", "Business Studies"],
    "Languages": ["English", "Afrikaans"],
    "Humanities": ["Geography", "History"],
}
SUBJECTS = [s for group in SUBJECT_GROUPS.values() for s in group]
SUBJECT_CHOICES = SUBJECTS + ["Other (tell us below)"]

# Special packages. "price" is shown big on the card, "detail" underneath.
# Use "price": None to show "Ask for a quote". Add, remove or rename freely.
SPECIALS = {
    "Exam Prep": {
        "for": "Grades 10-12 & first-year uni",
        "includes": ["Past papers", "Tricky topics", "Exam technique"],
        "sessions": "6 sessions",
        "price": "from R2 100",
        "detail": "Gr 10-11 R2 100 · Matric R2 400 · Uni R2 700",
        "totals": {"Grade 10-11": 2100, "Matric (Gr 12)": 2400, "First-year uni": 2700},
    },
    "NBT Prep": {
        "for": "Matrics applying to university",
        "includes": ["AQL", "MAT", "Timed practice"],
        "sessions": "4 sessions",
        "price": "R1 800",
        "detail": "2 × AQL + 2 × MAT, one-on-one",
        "totals": {"Grade 10-11": 1800, "Matric (Gr 12)": 1800},
    },
    "Holiday Catch-up": {
        "for": "Grades 1-12",
        "includes": ["Recap", "Next term preview", "Holiday hours"],
        "sessions": "5 sessions",
        "price": "from R950",
        "detail": "Gr 1-7 R950 · Gr 8-9 R1 200 · Gr 10-11 R1 650 · Matric R1 875",
        "totals": {"Primary (Gr 1-7)": 950, "Grade 8-9": 1200, "Grade 10-11": 1650, "Matric (Gr 12)": 1875},
    },
    "Varsity Head Start": {
        "for": "Matrics heading to university",
        "includes": ["First-year topics", "Study habits", "Insider tips"],
        "sessions": "4 sessions",
        "price": "R1 800",
        "detail": "One-on-one, before lectures start",
        "totals": {"Matric (Gr 12)": 1800, "First-year uni": 1800},
    },
}

# Tutor cards only appear once at least this many tutors are ticked "Show".
MIN_TUTORS_TO_SHOW = 3

# ---------- Data files ----------
DATA = Path(__file__).parent / "data"
DATA.mkdir(exist_ok=True)

SCHEMAS = {
    "bookings": {"Submitted": "text", "Parent": "text", "Phone": "text", "Email": "text",
                 "Student": "text", "Level": "text", "Subject": "text", "Package": "text",
                 "Total": "text", "Notes": "text"},
    "applications": {"Submitted": "text", "Name": "text", "Email": "text", "Phone": "text",
                     "University": "text", "Degree and year": "text", "Subjects": "text",
                     "Levels": "text", "Matric results": "text", "Uni transcript": "text",
                     "Availability": "text", "Why KeyStone": "text"},
    "packages": {"Start date": "text", "Student": "text", "Parent": "text", "Tutor": "text",
                  "Subject": "text", "Level": "text", "Package": "text", "Package size": "int",
                 "Price (R)": "int", "Lessons done": "int",
                 "Client paid": "bool", "Tutor paid": "bool"},
    "reviews": {"Quote": "text", "Who": "text", "Stars": "int", "Show": "bool",
                "Submitted": "text", "Name (private)": "text"},
    "products": {"ID": "text", "Title": "text", "Subject": "text", "Grade": "text", "Price": "int",
                 "Pages": "int", "Description": "text", "File": "text", "Previews": "text", "Show": "bool"},
    "orders": {"Submitted": "text", "Order ref": "text", "Product ID": "text", "Product": "text",
               "Amount": "int", "Name": "text", "Email": "text", "Phone": "text",
               "Proof of payment": "text", "Status": "text"},
    "tutors": {"Name": "text", "University": "text", "Degree": "text", "Subjects": "text",
               "Grades": "text", "Fun fact": "text", "Show": "bool"},
}


# ---------- Where data is stored ----------
# Online (Streamlit Cloud): a Google Sheet, one tab per table. Needs SHEET_URL and
# [gcp_service_account] in the app's secrets.
# On your laptop without those secrets: CSV files in the "data" folder.
def _secret(key):
    try:
        return st.secrets[key]
    except Exception:
        return None


USE_SHEETS = bool(_secret("SHEET_URL") and _secret("gcp_service_account"))


@st.cache_resource(show_spinner=False)
def _spreadsheet():
    import gspread
    from google.oauth2.service_account import Credentials

    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]),
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    return gspread.authorize(creds).open_by_url(st.secrets["SHEET_URL"])


def _worksheet(name: str):
    import gspread

    book = _spreadsheet()
    try:
        return book.worksheet(name)
    except gspread.WorksheetNotFound:
        ws = book.add_worksheet(title=name, rows=200, cols=len(SCHEMAS[name]))
        ws.update([list(SCHEMAS[name])], "A1", value_input_option="RAW")
        return ws


@st.cache_data(ttl=60, show_spinner=False)
def _read_raw(name: str) -> pd.DataFrame:
    if USE_SHEETS:
        values = _worksheet(name).get_all_values()
        if not values:
            return pd.DataFrame()
        header, rows = values[0], values[1:]
        rows = [r + [""] * (len(header) - len(r)) for r in rows if any(c.strip() for c in r)]
        return pd.DataFrame(rows, columns=header)
    path = DATA / f"{name}.csv"
    return pd.read_csv(path, dtype=str, keep_default_na=False) if path.exists() else pd.DataFrame()


def load(name: str) -> pd.DataFrame:
    """Read a table with tidy column types."""
    schema = SCHEMAS[name]
    df = _read_raw(name).copy()
    for col, kind in schema.items():
        if col not in df:
            df[col] = ""
        if kind == "int":
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)
        elif kind == "bool":
            df[col] = df[col].map(lambda v: str(v).strip().lower() in ("true", "1", "yes")).astype(bool)
        else:
            df[col] = df[col].fillna("").astype(str)
    return df[list(schema)].reset_index(drop=True)


def _cell(value, kind):
    """Plain Python values for Google Sheets (stored as-is, never run as formulas)."""
    if kind == "int":
        try:
            return int(value)
        except (TypeError, ValueError):
            return 0
    if kind == "bool":
        return bool(value) if not isinstance(value, str) else value.strip().lower() in ("true", "1", "yes")
    return "" if value is None or (isinstance(value, float) and pd.isna(value)) else str(value)


def _rows(name: str, df: pd.DataFrame) -> list:
    schema = SCHEMAS[name]
    return [[_cell(rec.get(c), k) for c, k in schema.items()] for rec in df.to_dict("records")]


def save(name: str, df: pd.DataFrame) -> None:
    if USE_SHEETS:
        ws = _worksheet(name)
        ws.clear()
        ws.update([list(SCHEMAS[name])] + _rows(name, df), "A1", value_input_option="RAW")
    else:
        df.to_csv(DATA / f"{name}.csv", index=False)
    _read_raw.clear()


def append(name: str, row: dict) -> None:
    if USE_SHEETS:
        ws = _worksheet(name)
        if ws.row_values(1) != list(SCHEMAS[name]):
            # Header missing or out of date: rewrite the tab with the right columns.
            save(name, pd.concat([load(name), pd.DataFrame([row])], ignore_index=True))
            return
        ws.append_rows(_rows(name, pd.DataFrame([row])), value_input_option="RAW")
        _read_raw.clear()
    else:
        save(name, pd.concat([load(name), pd.DataFrame([row])], ignore_index=True))


# First run on a laptop: start the tutor list with Palesa so the page isn't empty
if not USE_SHEETS and not (DATA / "tutors.csv").exists():
    save("tutors", pd.DataFrame([{
        "Name": "Palesa", "University": "", "Degree": "", "Subjects": "",
        "Grades": "", "Fun fact": "", "Show": True,
    }]))



# ---------- File storage (tutor transcripts) ----------
# Online: a private Supabase Storage bucket. Needs SUPABASE_URL and SUPABASE_KEY
# (the secret / service_role key) in the app's secrets. On your laptop without
# them: files go to data/uploads.
USE_SUPABASE = bool(_secret("SUPABASE_URL") and _secret("SUPABASE_KEY"))
BUCKET = _secret("SUPABASE_BUCKET") or "transcripts"
MAX_UPLOAD_MB = 10
UPLOAD_TYPES = ["pdf", "jpg", "jpeg", "png"]


def _sb_headers(extra=None):
    key = str(st.secrets["SUPABASE_KEY"])
    h = {"apikey": key, "Authorization": f"Bearer {key}"}
    h.update(extra or {})
    return h


def _safe(text: str) -> str:
    import re
    return re.sub(r"[^A-Za-z0-9]+", "-", text).strip("-").lower()[:40] or "file"


def store_bytes(data: bytes, path: str, content_type: str = "application/octet-stream") -> str:
    """Save raw bytes at `path` (e.g. 'workbooks/x.pdf') and return the path."""
    if USE_SUPABASE:
        import requests
        url = f"{str(st.secrets['SUPABASE_URL']).rstrip('/')}/storage/v1/object/{BUCKET}/{path}"
        r = requests.post(url, data=data, timeout=120, headers=_sb_headers(
            {"Content-Type": content_type, "x-upsert": "true"}))
        r.raise_for_status()
    else:
        target = DATA / "uploads" / path
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(data)
    return path


@st.cache_data(ttl=3600, show_spinner=False, max_entries=200)
def read_bytes(path: str) -> bytes:
    """Fetch a stored file's bytes (used for previews and downloads)."""
    if not path:
        return b""
    if USE_SUPABASE:
        import requests
        base = str(st.secrets["SUPABASE_URL"]).rstrip("/")
        r = requests.get(f"{base}/storage/v1/object/{BUCKET}/{path}", headers=_sb_headers(), timeout=120)
        r.raise_for_status()
        return r.content
    target = DATA / "uploads" / path
    return target.read_bytes() if target.exists() else b""


def store_file(uploaded, person: str, label: str, folder: str = "") -> str:
    """Save an uploaded file and return its storage path ('' if nothing uploaded)."""
    if uploaded is None:
        return ""
    ext = uploaded.name.rsplit(".", 1)[-1].lower() if "." in uploaded.name else "bin"
    path = f"{datetime.now():%Y%m%d-%H%M%S}_{_safe(person)}_{label}.{ext}"
    if folder:
        path = f"{folder}/{path}"
    return store_bytes(uploaded.getvalue(), path, uploaded.type or "application/octet-stream")


def file_link(path: str) -> str:
    """A temporary (1 hour) link to view a stored file, for the admin page."""
    if not path:
        return ""
    if USE_SUPABASE:
        import requests
        base = str(st.secrets["SUPABASE_URL"]).rstrip("/")
        r = requests.post(f"{base}/storage/v1/object/sign/{BUCKET}/{path}", json={"expiresIn": 3600},
                          headers=_sb_headers(), timeout=30)
        if not r.ok:
            return ""
        signed = r.json().get("signedURL") or r.json().get("signedUrl") or ""
        return f"{base}/storage/v1{signed}" if signed.startswith("/") else signed
    return ""

def admin_password() -> str | None:
    """The admin password comes ONLY from Streamlit secrets. No secret = admin stays locked."""
    try:
        pw = str(st.secrets["ADMIN_PASSWORD"]).strip()
        return pw or None
    except Exception:
        return None


# ---------- Shop helpers ----------
# Banking details come from the app's secrets (BANK_NAME, ACCOUNT_NAME, ACCOUNT_NUMBER,
# BRANCH_CODE, ACCOUNT_TYPE) so they live in one place and are easy to change.
def bank_details() -> dict:
    keys = {"Bank": "BANK_NAME", "Account name": "ACCOUNT_NAME", "Account number": "ACCOUNT_NUMBER",
            "Branch code": "BRANCH_CODE", "Account type": "ACCOUNT_TYPE"}
    return {label: str(_secret(k)) for label, k in keys.items() if _secret(k)}


def new_order_ref() -> str:
    import secrets as _s
    alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"   # no 0/O/1/I mix-ups
    return "KS-" + "".join(_s.choice(alphabet) for _ in range(5))


def pdf_previews(pdf: bytes, pages: int = 3, zoom: float = 1.3):
    """PNG images of the first few pages of a PDF, plus its total page count."""
    import pymupdf as fitz
    doc = fitz.open(stream=pdf, filetype="pdf")
    images = [doc[i].get_pixmap(matrix=fitz.Matrix(zoom, zoom)).tobytes("png")
              for i in range(min(pages, len(doc)))]
    return images, len(doc)


def order_links(title: str, price: int):
    """WhatsApp and email links with the order message already filled in."""
    msg = f"Hi KeyStone! I'd like to order the {title} ({rand(price)}). Please send me the payment details."
    wa = f"https://wa.me/{WHATSAPP}?text={quote(msg)}"
    mail = f"mailto:{EMAIL}?subject={quote('Workbook order: ' + title)}&body={quote(msg)}"
    return wa, mail


def stamp_pdf(pdf: bytes, buyer: str, ref: str) -> bytes:
    """Add a small 'Licensed to ...' line to the bottom of every page."""
    import pymupdf as fitz
    doc = fitz.open(stream=pdf, filetype="pdf")
    text = f"Licensed to {buyer} · Order {ref} · For personal use only. Not for sharing or resale."
    for page in doc:
        r = page.rect
        page.insert_textbox(fitz.Rect(36, r.height - 26, r.width - 36, r.height - 8), text,
                            fontsize=7.5, fontname="helv", color=(0.42, 0.39, 0.35), align=1)
    return doc.tobytes(garbage=3, deflate=True)


def rand(n: int) -> str:
    return f"R{n:,}".replace(",", " ")


def package_value(level: str, size: int) -> int:
    single, four, ten = PRICES.get(level, (0, 0, 0))
    per = ten if size >= 10 else four if size >= 4 else single
    return per * size


def booking_total(level: str, package: str):
    """(total in rand or None, short explanation) for a level + package choice."""
    if package in PACK_SIZES:
        size = PACK_SIZES[package]
        total = package_value(level, size)
        per = total // size
        return total, ("1 lesson" if size == 1 else f"{size} lessons at {rand(per)} each")
    if package in SPECIALS:
        sp = SPECIALS[package]
        total = sp.get("totals", {}).get(level)
        if total:
            return total, sp.get("sessions", "")
        return None, f"{package} isn't usually offered for {level}. We'll chat about the best option."
    return None, "No stress, we'll help you choose the right package."


# ---------- Mascot (same character as the Instagram posts) ----------
MASCOT = """
<svg viewBox="0 0 370 380" width="100%" style="max-width:320px" role="img" aria-label="KeyStone mascot waving">
<g stroke="#2A2622" stroke-width="5" stroke-linecap="round" stroke-linejoin="round">
<path d="M150 258 Q104 248 90 204" fill="none"/><circle cx="86" cy="194" r="14" fill="#F6D97A"/>
<path d="M200 258 Q236 264 256 250" fill="none"/>
<rect x="252" y="214" width="46" height="60" rx="7" fill="#FFFFFF"/><rect x="252" y="234" width="46" height="20" fill="#F2B8A0"/><rect x="246" y="204" width="58" height="14" rx="6" fill="#FFFFFF"/>
<path d="M266 192 Q258 180 266 168" fill="none" stroke-width="4"/><path d="M284 192 Q276 180 284 168" fill="none" stroke-width="4"/>
<circle cx="256" cy="252" r="14" fill="#F6D97A"/>
<path d="M162 340 L156 360" fill="none"/><path d="M190 340 L196 360" fill="none"/>
<ellipse cx="148" cy="362" rx="18" ry="9" fill="#2A2622"/><ellipse cx="204" cy="362" rx="18" ry="9" fill="#2A2622"/>
<rect x="196" y="282" width="38" height="20" rx="5" fill="#F6D97A"/><rect x="196" y="312" width="26" height="20" rx="5" fill="#F6D97A"/>
<rect x="145" y="200" width="60" height="146" rx="18" fill="#F6D97A"/><circle cx="175" cy="135" r="100" fill="#F6D97A"/>
<circle cx="175" cy="135" r="82" fill="none" stroke-width="3" stroke-dasharray="2 11" opacity="0.4"/><circle cx="175" cy="62" r="14" fill="#FBF6EC"/>
</g>
<ellipse cx="150" cy="140" rx="10" ry="13" fill="#2A2622"/><ellipse cx="200" cy="140" rx="10" ry="13" fill="#2A2622"/>
<circle cx="153" cy="135" r="3.5" fill="#FFFFFF"/><circle cx="203" cy="135" r="3.5" fill="#FFFFFF"/>
<ellipse cx="124" cy="168" rx="14" ry="8" fill="#F2B8A0"/><ellipse cx="226" cy="168" rx="14" ry="8" fill="#F2B8A0"/>
<path d="M156 168 Q175 192 194 168" fill="none" stroke="#2A2622" stroke-width="5" stroke-linecap="round"/>
</svg>
"""

# ---------- Page setup & brand styling ----------
st.set_page_config(page_title="KeyStone Tutors", page_icon="🔑", layout="centered")

st.markdown(
    """
    <link href="https://fonts.googleapis.com/css2?family=Caveat:wght@700&family=DM+Sans:wght@400;500;700&family=Fraunces:ital,opsz,wght@0,9..144,700;1,9..144,700&family=Space+Mono:wght@700&display=swap" rel="stylesheet">
    <style>
      html, body, [class*="css"], .stMarkdown, p, li, label { font-family: 'DM Sans', sans-serif; }
      h1, h2, h3 { font-family: 'Fraunces', Georgia, serif !important; letter-spacing: -0.02em; }
      .kicker { font-family: 'Space Mono', monospace; font-weight: 700; letter-spacing: 0.24em; font-size: 0.85rem; }
      .hero { font-family: 'Fraunces', Georgia, serif; font-weight: 700; font-size: 3.6rem; line-height: 0.95; margin: 0.4rem 0 1rem; }
      .hl { font-style: italic; background: linear-gradient(transparent 58%, #F6D97A 58%, #F6D97A 92%, transparent 92%); padding: 0 6px; }
      .hand { font-family: 'Caveat', cursive; font-size: 2rem; color: #D9705A; transform: rotate(-2deg); display: inline-block; }
      .card { background: #FFFDF8; border: 3px solid #2A2622; border-radius: 18px; padding: 1.2rem 1.4rem; box-shadow: 6px 8px 0 rgba(42,38,34,0.12); height: 100%; margin-bottom: 0.9rem; }
      .card h4 { font-family: 'Fraunces', Georgia, serif; margin: 0 0 0.4rem; font-size: 1.3rem; }
      .meta { font-family: 'Space Mono', monospace; font-weight: 700; font-size: 0.75rem; letter-spacing: 0.1em; color: #8A8177; text-transform: uppercase; }
      .step { font-family: 'Space Mono', monospace; font-weight: 700; color: #D9705A; }
      .quote { font-family: 'Caveat', cursive; font-size: 1.7rem; line-height: 1.15; }
      .stars { color: #D9705A; letter-spacing: 0.15em; }
      .result { background: #F6D97A; border: 3px solid #2A2622; border-radius: 18px; padding: 1.2rem 1.4rem; }
      .tile { background: #FFFDF8; border: 3px solid #2A2622; border-radius: 18px; padding: 1.1rem 1rem; text-align: center; box-shadow: 6px 8px 0 rgba(42,38,34,0.12); height: 100%; margin-bottom: 0.9rem; }
      .tile .ic { width: 56px; height: 56px; border-radius: 50%; background: #F6D97A; border: 3px solid #2A2622; display: flex; align-items: center; justify-content: center; margin: 0 auto 0.6rem; }
      .tile .tt { font-family: 'Fraunces', Georgia, serif; font-weight: 700; font-size: 1.15rem; line-height: 1.15; }
      .tile .ts { font-size: 0.9rem; color: #6B635A; margin-top: 0.25rem; }
      .chips { display: flex; flex-wrap: wrap; gap: 0.5rem; margin: 0.4rem 0 1rem; }
      .chip { background: #FFFDF8; border: 2px solid #2A2622; border-radius: 999px; padding: 0.3rem 0.85rem; font-weight: 500; font-size: 0.95rem; }
      .chip.y { background: #F6D97A; }
      .steps { display: flex; gap: 0.8rem; flex-wrap: wrap; margin: 0.5rem 0 1rem; }
      .stp { flex: 1 1 110px; text-align: center; }
      .stp .n { width: 48px; height: 48px; border-radius: 50%; background: #2A2622; color: #F6D97A; font-family: 'Fraunces', Georgia, serif; font-weight: 700; font-size: 1.4rem; display: flex; align-items: center; justify-content: center; margin: 0 auto 0.5rem; }
      .stp .tt { font-family: 'Fraunces', Georgia, serif; font-weight: 700; font-size: 1.1rem; }
      .stp .ts { font-size: 0.88rem; color: #6B635A; }
      .price { background: #FFFDF8; border: 3px solid #2A2622; border-radius: 18px; padding: 1rem; text-align: center; box-shadow: 6px 8px 0 rgba(42,38,34,0.12); margin-bottom: 0.9rem; }
      .price .lv { font-family: 'Space Mono', monospace; font-weight: 700; font-size: 0.72rem; letter-spacing: 0.08em; text-transform: uppercase; color: #6B635A; }
      .price .big { font-family: 'Fraunces', Georgia, serif; font-weight: 700; font-size: 1.9rem; line-height: 1.1; margin: 0.3rem 0 0; white-space: nowrap; }
      .price-grid { display: grid; grid-template-columns: repeat(5, minmax(0, 1fr)); gap: 0.75rem; margin-bottom: 1rem; }
      .price-grid .price { margin-bottom: 0; padding: 0.9rem 0.6rem; }
      .sp-grid { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 0.75rem; margin-bottom: 1rem; }
      .sp-grid .price { margin-bottom: 0; padding: 0.9rem 0.6rem; }
      .sp-grid .price .big { font-size: 1.35rem; }
      .price-grid .price .big { font-size: 1.65rem; }
      @media (max-width: 760px) {
        .price-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
        .price-grid .price:last-child { grid-column: span 2; }
        .price-grid .price .big { font-size: 1.9rem; }
        .sp-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      }
      .price .per { font-size: 0.8rem; color: #6B635A; margin-bottom: 0.5rem; }
      .price .row { display: flex; justify-content: space-between; gap: 0.5rem; white-space: nowrap; font-size: 0.85rem; border-top: 1px dashed #CFC6B8; padding-top: 0.3rem; margin-top: 0.3rem; }
      .sp-price { display: flex; align-items: center; gap: 0.6rem; flex-wrap: wrap; margin-top: 0.2rem; }
      .sp-price .big { font-family: 'Fraunces', Georgia, serif; font-weight: 700; font-size: 1.7rem; white-space: nowrap; }
      .card .per { font-size: 0.82rem; color: #6B635A; margin-top: 0.3rem; }
      .bank { background: #FFFDF8; border: 3px dashed #2A2622; border-radius: 18px; padding: 1rem 1.2rem; margin: 0.5rem 0 1rem; }
      .bank .r { display: flex; justify-content: space-between; gap: 1rem; padding: 0.25rem 0; border-bottom: 1px dashed #CFC6B8; }
      .bank .r:last-child { border-bottom: 0; }
      .bank .r b { font-family: 'Space Mono', monospace; }
      .ref { display: inline-block; background: #F6D97A; border: 3px solid #2A2622; border-radius: 12px; padding: 0.3rem 0.9rem; font-family: 'Space Mono', monospace; font-weight: 700; font-size: 1.4rem; letter-spacing: 0.08em; }
      .sub { font-size: 1.15rem; color: #2A2622; margin: 0 0 1rem; }
      h3.sec { font-family: 'Fraunces', Georgia, serif; margin: 1.6rem 0 0.4rem; }
      .foot { font-family: 'Space Mono', monospace; font-weight: 700; font-size: 0.8rem; letter-spacing: 0.08em; color: #6B635A; text-align: center; margin-top: 3rem; }
      @media (max-width: 640px) { .hero { font-size: 2.6rem; } }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="kicker">KEYSTONE TUTORS</div>', unsafe_allow_html=True)



# ---------- Small visual helpers ----------
_ICON_PATHS = {
    "person": '<circle cx="12" cy="8" r="4"/><path d="M4 21c0-4 4-6 8-6s8 2 8 6"/>',
    "laptop": '<rect x="3" y="5" width="18" height="12" rx="2"/><path d="M1 20h22"/>',
    "match": '<circle cx="8" cy="9" r="3"/><circle cx="16" cy="9" r="3"/><path d="M2 20c0-3 3-5 6-5M22 20c0-3-3-5-6-5"/>',
    "star": '<path d="M12 3l2.8 5.7 6.2.9-4.5 4.4 1 6.2L12 17.3 6.5 20.2l1-6.2L3 9.6l6.2-.9z"/>',
    "clock": '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
    "wallet": '<rect x="3" y="6" width="18" height="13" rx="2"/><path d="M3 10h18M16 14h2"/>',
    "shield": '<path d="M12 3l8 3v6c0 5-3.5 8-8 9-4.5-1-8-4-8-9V6z"/><path d="M8.5 12l2.5 2.5 4.5-5"/>',
    "chat": '<path d="M21 12a8 8 0 0 1-11.6 7.1L4 20l1-4.6A8 8 0 1 1 21 12z"/>',
    "briefcase": '<rect x="3" y="7" width="18" height="13" rx="2"/><path d="M9 7V5a2 2 0 0 1 2-2h2a2 2 0 0 1 2 2v2"/>',
    "target": '<circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="5"/><circle cx="12" cy="12" r="1"/>',
    "search": '<circle cx="11" cy="11" r="7"/><path d="M21 21l-5-5"/>',
    "book": '<path d="M4 5a2 2 0 0 1 2-2h12v16H6a2 2 0 0 0-2 2z"/><path d="M4 21a2 2 0 0 1 2-2h12"/>',
}


def icon(name: str) -> str:
    return ('<svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="#2A2622" '
            'stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">'
            f"{_ICON_PATHS[name]}</svg>")


def tiles(items, per_row=3):
    """Icon tiles: items = [(icon, title, short line), ...]"""
    for start in range(0, len(items), per_row):
        cols = st.columns(per_row)
        for col, (ic, title, sub) in zip(cols, items[start:start + per_row]):
            col.markdown(f'<div class="tile"><div class="ic">{icon(ic)}</div><div class="tt">{title}</div>'
                         f'<div class="ts">{sub}</div></div>', unsafe_allow_html=True)


def chips(words, highlight=()):
    html = "".join(f'<span class="chip{" y" if w in highlight else ""}">{w}</span>' for w in words)
    st.markdown(f'<div class="chips">{html}</div>', unsafe_allow_html=True)


def steps(items):
    html = "".join(f'<div class="stp"><div class="n">{i}</div><div class="tt">{t}</div><div class="ts">{d}</div></div>'
                   for i, (t, d) in enumerate(items, 1))
    st.markdown(f'<div class="steps">{html}</div>', unsafe_allow_html=True)


def section(title):
    st.markdown(f'<h3 class="sec">{title}</h3>', unsafe_allow_html=True)


def go(label, page, primary=False, key=None):
    if st.button(label, type="primary" if primary else "secondary", key=key or f"go_{label}",
                 width="stretch"):
        st.switch_page(page)


# ---------- Home ----------
def page_home():
    left, right = st.columns([3, 2], vertical_alignment="center")
    with left:
        st.markdown(
            '<div class="hero">Tutoring that actually<br><span class="hl">clicks.</span></div>'
            '<div class="hand">run by students, for students</div>',
            unsafe_allow_html=True,
        )
        chips(["One-on-one", "Online", "Gr 1 to first-year uni", "IEB & NSC"])
        b1, b2 = st.columns(2)
        with b1:
            go("Book a lesson", P_BOOK, primary=True, key="home_book")
        with b2:
            go("See prices", P_PRICING, key="home_prices")
    with right:
        st.markdown(MASCOT, unsafe_allow_html=True)

    section("Why KeyStone")
    tiles([
        ("person", "One-on-one", "Just your child and their tutor"),
        ("laptop", "Online", "Learn from home"),
        ("match", "Matched tutors", "The right fit for each student"),
    ])

    section("Subjects")
    chips(SUBJECTS)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="tile"><div class="ic">' + icon("target") + '</div>'
                    '<div class="tt">Exams or NBTs coming up?</div></div>', unsafe_allow_html=True)
        go("Special packages", P_SPECIAL, key="home_special")
    with c2:
        st.markdown('<div class="tile"><div class="ic">' + icon("briefcase") + '</div>'
                    '<div class="tt">Uni student? Tutor with us</div></div>', unsafe_allow_html=True)
        go("Join the team", P_JOIN, key="home_join")

    reviews = load("reviews")
    reviews = reviews[reviews["Show"] & (reviews["Quote"].str.strip() != "")]
    if not reviews.empty:
        section("What parents say")
        cols = st.columns(2)
        for i, (_, r) in enumerate(reviews.iterrows()):
            stars = "★" * max(0, min(5, int(r["Stars"])))
            cols[i % 2].markdown(
                f'<div class="card"><div class="stars">{stars}</div>'
                f'<div class="quote">"{r["Quote"]}"</div><div class="meta">{r["Who"]}</div></div>',
                unsafe_allow_html=True,
            )

# ---------- Our tutors ----------
def page_tutors():
    st.markdown('<div class="hero" style="font-size:2.8rem">Real students.<br>'
                '<span class="hl">Real marks.</span></div>', unsafe_allow_html=True)
    st.markdown('<p class="sub">Uni students who remember what made it click.</p>', unsafe_allow_html=True)
    tiles([
        ("star", "Top of their class", "In the subjects they teach"),
        ("chat", "Great explainers", "Patient, clear, relatable"),
        ("match", "Matched to you", "Subject, grade and style"),
        ("shield", "Managed by us", "We stay in the loop"),
    ], per_row=2)

    tutors = load("tutors")
    tutors = tutors[tutors["Show"] & (tutors["Name"].str.strip() != "")]
    if len(tutors) >= MIN_TUTORS_TO_SHOW:
        section("Meet the team")
        cols = st.columns(2)
        for i, (_, t) in enumerate(tutors.iterrows()):
            study = " · ".join(x for x in [t["Degree"], t["University"]] if x.strip())
            lines = [f"<h4>{t['Name']}</h4>"]
            if study:
                lines.append(f'<div class="meta">{study}</div>')
            if t["Subjects"].strip():
                lines.append(f"<p>{t['Subjects']}</p>")
            if t["Fun fact"].strip():
                lines.append(f'<div class="hand" style="font-size:1.4rem">{t["Fun fact"]}</div>')
            cols[i % 2].markdown(f'<div class="card">{"".join(lines)}</div>', unsafe_allow_html=True)
    go("Book a lesson", P_BOOK, primary=True, key="tutors_book")

# ---------- Pricing ----------
def page_pricing():
    st.markdown('<div class="hero" style="font-size:2.8rem">Simple <span class="hl">pricing.</span></div>',
                unsafe_allow_html=True)
    st.markdown('<p class="sub">Per session. Bigger pack, better price.</p>', unsafe_allow_html=True)
    short = {"Primary (Gr 1-7)": "Gr 1-7", "Grade 8-9": "Gr 8-9", "Grade 10-11": "Gr 10-11",
             "Matric (Gr 12)": "Matric", "First-year uni": "Uni"}
    cards = "".join(
        f'<div class="price"><div class="lv">{short.get(level, level)}</div>'
        f'<div class="big">{rand(ten)}</div><div class="per">with a 10-pack</div>'
        f'<div class="row"><span>Single</span><b>{rand(single)}</b></div>'
        f'<div class="row"><span>4-pack</span><b>{rand(four)}</b></div></div>'
        for level, (single, four, ten) in PRICES.items()
    )
    st.markdown(f'<div class="price-grid">{cards}</div>', unsafe_allow_html=True)

    section("Special packages")
    sp_cards = "".join(
        f'<div class="price"><div class="lv">{name}</div>'
        f'<div class="big">{sp.get("price") or "Quote"}</div><div class="per">{sp.get("sessions", "")}</div></div>'
        for name, sp in SPECIALS.items()
    )
    st.markdown(f'<div class="sp-grid">{sp_cards}</div>', unsafe_allow_html=True)

    section("Your total")
    q1, q2 = st.columns(2)
    level = q1.selectbox("Level", LEVELS, key="q_level")
    pack = q2.selectbox("Package", list(PACK_SIZES), key="q_pack")
    size = PACK_SIZES[pack]
    total = package_value(level, size)
    st.markdown(
        f'<div class="price" style="text-align:left"><div class="lv">Package total</div>'
        f'<div class="big">{rand(total)}</div><div class="per">{rand(total // size)} per session</div></div>',
        unsafe_allow_html=True,
    )
    b1, b2 = st.columns(2)
    with b1:
        go("Book a lesson", P_BOOK, primary=True, key="price_book")
    with b2:
        go("Exam & NBT prep", P_SPECIAL, key="price_special")

# ---------- Special packages ----------
def page_special():
    st.markdown('<div class="hero" style="font-size:2.8rem">Special <span class="hl">packages.</span></div>',
                unsafe_allow_html=True)
    st.markdown('<p class="sub">For the moments that matter most.</p>', unsafe_allow_html=True)
    scols = st.columns(2)
    for i, (name, sp) in enumerate(SPECIALS.items()):
        price = sp.get("price") or "Ask for a quote"
        items = "".join(f'<span class="chip">{x}</span>' for x in sp["includes"])
        scols[i % 2].markdown(
            f'<div class="card"><div class="meta">{sp["for"]}</div><h4>{name}</h4>'
            f'<div class="chips">{items}</div>'
            f'<div class="sp-price"><span class="big">{price}</span>'
            f'<span class="chip y">{sp.get("sessions", "")}</span></div>'
            f'<div class="per">{sp.get("detail", "")}</div></div>',
            unsafe_allow_html=True,
        )
    go("Book a lesson", P_BOOK, primary=True, key="special_book")

# ---------- Which package? ----------
def page_quiz():
    st.markdown('<div class="hero" style="font-size:2.6rem">Which package <span class="hl">fits?</span></div>',
                unsafe_allow_html=True)
    st.markdown('<p class="sub">3 quick questions.</p>', unsafe_allow_html=True)
    qa = st.selectbox("1. What level is your child in?", LEVELS, key="quiz_level")
    qb = st.radio(
        "2. When is their next big test or exam?",
        ["Within the next 2 weeks", "In about a month", "More than a month away"],
        key="quiz_when",
    )
    qc = st.radio(
        "3. What kind of help do they need?",
        ["Help with one tricky topic", "Steady support to keep up in class",
         "A big push to bring their marks up", "Getting ready for the NBTs"],
        key="quiz_need",
    )

    senior = qa in ("Grade 10-11", "Matric (Gr 12)", "First-year uni")
    score = {"Within the next 2 weeks": 0, "In about a month": 1, "More than a month away": 2}[qb]
    score += {"Help with one tricky topic": 0, "Steady support to keep up in class": 1,
              "A big push to bring their marks up": 2, "Getting ready for the NBTs": 2}[qc]
    if qc == "Getting ready for the NBTs":
        rec, why = "NBT Prep", "Practice on both NBT papers, with timed runs so test day feels familiar."
    elif qc == "A big push to bring their marks up" and qb != "More than a month away" and senior:
        rec, why = "Exam Prep", "With exams this close, a focused run of past papers and exam technique works best."
    elif qc == "Help with one tricky topic" and score <= 1:
        rec, why = "Single lesson", "A focused session or two on that one topic should do the trick."
    elif score >= 3:
        rec, why = "10-lesson pack", ("There's time to build real understanding, and a 10-pack gives "
                                      "you the lowest price per lesson.")
    else:
        rec, why = "4-lesson pack", "Enough sessions to see real progress, at a lower price per lesson."

    if rec in SPECIALS:
        sp = SPECIALS[rec]
        cost = (f"<b>{sp['price']}</b> for {sp.get('sessions', '')}. {sp.get('detail', '')}"
                if sp.get("price") else "Ask us for a quote.")
    else:
        size = PACK_SIZES[rec]
        total = package_value(qa, size)
        cost = f"<b>{rand(total)}</b> in total ({rand(total // size)} per session for {qa})."
    st.markdown(
        f'<div class="result"><div class="meta" style="color:#2A2622">We recommend</div>'
        f'<h3 style="margin:0.2rem 0">{rec}</h3><p>{why}</p>{cost}</div>',
        unsafe_allow_html=True,
    )
    go("Book a lesson", P_BOOK, primary=True, key="quiz_book")

# ---------- How it works ----------
def page_how():
    st.markdown('<div class="hero" style="font-size:2.8rem">How it <span class="hl">works.</span></div>',
                unsafe_allow_html=True)
    steps([
        ("Book", "Fill in a quick form"),
        ("Get matched", "We pick the right tutor"),
        ("Choose a pack", "Single, 4 or 10 lessons"),
        ("Learn", "Online, your schedule"),
    ])
    go("Book a lesson", P_BOOK, primary=True, key="how_book")

# ---------- FAQ ----------
def page_faq():
    st.markdown('<div class="hero" style="font-size:2.8rem">Questions?</div>', unsafe_allow_html=True)
    faqs = [
        ("Are the lessons online or in person?",
         "Online, one-on-one. Your child learns from home with their own tutor."),
        ("Which grades and syllabuses do you cover?",
         "Grades 1 to 12 (IEB and NSC) and first-year university."),
        ("Which subjects do you offer?",
         ", ".join(SUBJECTS[:-1]) + f" and {SUBJECTS[-1]}. Need something else? Ask us and we'll do our best to match you."),
        ("Do you help with exam or NBT preparation?",
         "Yes. Have a look at the Special packages page for Exam Prep, NBT Prep, Holiday Catch-up "
         "and Varsity Head Start."),
        ("How do the packs work?",
         "You choose a single lesson, a 4-pack or a 10-pack. The bigger the pack, the less each lesson "
         "costs. Your tutor works through the lessons with your child and you can top up any time."),
        ("How do we arrange session times?",
         "Once you're matched, we create a WhatsApp group with you, the tutor and KeyStone. "
         "Your tutor shares session times there."),
        ("What if we need to miss a session?",
         "Let your tutor know in the group chat as early as you can and they'll find a new time with you."),
        ("Can I share a workbook I bought?",
         "Workbooks are for the buyer's own use. Each copy is licensed to the person who bought it, "
         "and their name and order number appear on every page."),
        ("Who are the tutors?",
         "University students who did well in the subjects they teach. Read more on the Our tutors page."),
        ("How do I book?",
         f"Use the Book a lesson page, WhatsApp us on {PHONE_DISPLAY}, or email {EMAIL}."),
    ]
    for q, a in faqs:
        with st.expander(q):
            st.write(a)

# ---------- Book ----------
def page_book():
    st.markdown('<div class="hero" style="font-size:2.8rem">Book a <span class="hl">lesson.</span></div>',
                unsafe_allow_html=True)
    st.markdown('<p class="sub">We\'ll WhatsApp you to set up the first session.</p>', unsafe_allow_html=True)
    done = st.session_state.pop("booked_ok", None)
    if done:
        st.success(done)
        st.balloons()

    n = st.session_state.setdefault("book_form_n", 0)   # bumping n gives a fresh, empty form

    def k(name):
        return f"book{n}_{name}"

    with st.container(border=True):
        parent = st.text_input("Parent / guardian name", key=k("parent"))
        c1, c2 = st.columns(2)
        phone = c1.text_input("WhatsApp number", key=k("phone"))
        email = c2.text_input("Email (optional)", key=k("email"))
        student = st.text_input("Student name", key=k("student"))
        b1, b2 = st.columns(2)
        grade = b1.selectbox("Level", LEVELS, key=k("level"))
        subject = b2.selectbox("Subject", SUBJECT_CHOICES, key=k("subject"))
        package = st.selectbox("Package", list(PACK_SIZES) + list(SPECIALS) + ["Not sure yet"], key=k("package"))

        total, note = booking_total(grade, package)
        st.markdown(
            f'<div class="price" style="text-align:left"><div class="lv">Your total</div>'
            f'<div class="big">{rand(total) if total else "To be confirmed"}</div>'
            f'<div class="per">{note}</div></div>',
            unsafe_allow_html=True,
        )
        notes = st.text_area("Anything we should know? (topics, exam dates, availability)", key=k("notes"))
        submitted = st.button("Book now", type="primary", key=k("submit"))

    if submitted:
        if not parent.strip() or not student.strip() or not phone.strip():
            st.error("Please add your name, your WhatsApp number and the student's name.")
        else:
            append("bookings", {
                "Submitted": datetime.now().strftime("%Y-%m-%d %H:%M"), "Parent": parent,
                "Phone": phone, "Email": email, "Student": student, "Level": grade,
                "Subject": subject, "Package": package,
                "Total": rand(total) if total else "To be confirmed", "Notes": notes,
            })
            st.session_state["book_form_n"] += 1
            st.session_state["booked_ok"] = (
                f"Thank you, {parent.strip().split()[0]}! We've received your booking for "
                f"{student.strip()} and will be in touch on WhatsApp soon."
            )
            st.rerun()
    st.caption(f"Questions first? WhatsApp us on {PHONE_DISPLAY} or email {EMAIL}.")


# ---------- Join our team ----------
def page_join():
    st.markdown(
        '<div class="kicker" style="color:#D9705A">WE\'RE HIRING TUTORS</div>'
        '<div class="hero" style="font-size:2.8rem">Earn while you study.<br>'
        '<span class="hl">Teach what you\'re good at.</span></div>',
        unsafe_allow_html=True,
    )
    tiles([
        ("wallet", "Bigger share", "Most of every session fee"),
        ("search", "We find students", "No chasing parents"),
        ("clock", "Your hours", "Fits your timetable"),
        ("laptop", "100% online", "Teach from anywhere"),
        ("shield", "No admin", "Bookings & payments sorted"),
        ("briefcase", "CV boost", "Real teaching experience"),
    ])
    section("You'll fit in if you're")
    chips(["At uni or college", "Strong in your subjects", "Patient", "Reliable"], highlight=())
    section("How it works")
    steps([("Apply", "5-minute form"), ("Chat", "We get to know you"), ("Get matched", "Your subjects & grades"),
           ("Teach", "Online sessions"), ("Get paid", "After each package")])

    section("Apply now")
    if st.session_state.pop("applied_ok", False):
        st.success("Thanks for applying! We'll be in touch on WhatsApp or email.")
        st.balloons()
    with st.form(f"apply_{st.session_state.setdefault('apply_form_n', 0)}"):
        a1, a2 = st.columns(2)
        name = a1.text_input("Full name")
        phone = a2.text_input("WhatsApp number")
        email = st.text_input("Email")
        a3, a4 = st.columns(2)
        uni = a3.text_input("University")
        degree = a4.text_input("Degree and year (e.g. BSc Eng, 3rd year)")
        subj = st.multiselect("Subjects you can tutor", SUBJECTS + ["Other"])
        lvls = st.multiselect("Levels you can tutor", LEVELS)
        st.markdown("**Upload your results** (PDF or photo, max 10 MB each)")
        u1, u2 = st.columns(2)
        matric_file = u1.file_uploader("Matric results", type=UPLOAD_TYPES,
                                       help="Needed if you want to teach any school grade")
        uni_file = u2.file_uploader("University transcript", type=UPLOAD_TYPES,
                                    help="Needed if you want to teach first-year uni")
        avail = st.text_input("When are you usually free? (e.g. weekday evenings, Saturdays)")
        why = st.text_area("Why would you like to tutor with KeyStone?")
        applied = st.form_submit_button("Send my application", type="primary")

    if applied:
        school = any(l != "First-year uni" for l in lvls)
        varsity = "First-year uni" in lvls
        problems = []
        if not name.strip() or not (phone.strip() or email.strip()):
            problems.append("your name and a way to contact you")
        if not lvls:
            problems.append("the levels you can tutor")
        if school and matric_file is None:
            problems.append("your matric results (you chose school grades)")
        if varsity and uni_file is None:
            problems.append("your university transcript (you chose first-year uni)")
        too_big = [f.name for f in (matric_file, uni_file) if f is not None and f.size > MAX_UPLOAD_MB * 1024 * 1024]
        if problems:
            st.error("Please add " + "; ".join(problems) + ".")
        elif too_big:
            st.error(f"{', '.join(too_big)} is over {MAX_UPLOAD_MB} MB. Please upload a smaller file or photo.")
        else:
            try:
                with st.spinner("Uploading your documents..."):
                    matric_path = store_file(matric_file, name, "matric")
                    uni_path = store_file(uni_file, name, "uni")
            except Exception:
                st.error("Sorry, your documents couldn't be uploaded just now. Please try again in a minute, "
                         f"or WhatsApp us on {PHONE_DISPLAY}.")
            else:
                append("applications", {
                    "Submitted": datetime.now().strftime("%Y-%m-%d %H:%M"), "Name": name, "Email": email,
                    "Phone": phone, "University": uni, "Degree and year": degree,
                    "Subjects": ", ".join(subj), "Levels": ", ".join(lvls),
                    "Matric results": matric_path, "Uni transcript": uni_path,
                    "Availability": avail, "Why KeyStone": why,
                })
                st.session_state["apply_form_n"] += 1   # fresh, empty form
                st.session_state["applied_ok"] = True
                st.rerun()

# ---------- Shop ----------
def page_shop():
    st.markdown('<div class="hero" style="font-size:2.8rem">Workbooks & <span class="hl">study guides.</span></div>',
                unsafe_allow_html=True)
    st.markdown('<p class="sub">Written by tutors. Worked solutions explained like a lesson.</p>',
                unsafe_allow_html=True)
    products = load("products")
    products = products[products["Show"] & (products["Title"].str.strip() != "")]
    if products.empty:
        st.info("Our first workbooks are on their way. Check back soon!")
        return
    steps([("Preview", "See the first 3 pages"), ("Pay by EFT", "Use your order reference"),
           ("Upload proof", "Takes a minute"), ("Download", "Confirmed within 24 hours")])

    for _, p in products.iterrows():
        with st.container(border=True):
            meta = " · ".join(x for x in [p["Grade"], p["Subject"]] if x.strip())
            pages_txt = f" · {p['Pages']} pages" if p["Pages"] else ""
            st.markdown(
                f'<div class="meta">{meta}{pages_txt}</div><h3 style="margin:0.2rem 0">{p["Title"]}</h3>'
                f'<p style="margin:0 0 0.4rem">{p["Description"]}</p>'
                f'<span class="chip y"><b>{rand(int(p["Price"]))}</b></span>',
                unsafe_allow_html=True,
            )
            previews = [x for x in p["Previews"].split("|") if x.strip()]
            if previews:
                with st.expander("Preview the first pages", expanded=True):
                    cols = st.columns(len(previews))
                    for col, path in zip(cols, previews):
                        try:
                            col.image(read_bytes(path), width="stretch")
                        except Exception:
                            col.caption("Preview unavailable")
            wa, mail = order_links(p["Title"], int(p["Price"]))
            b1, b2, b3 = st.columns(3)
            if b1.button("Buy now (EFT)", key=f"buy_{p['ID']}", type="primary", width="stretch"):
                st.session_state["checkout"] = p["ID"]
                st.session_state.pop("order_ref", None)
                st.switch_page(P_CHECKOUT)
            b2.link_button("Order on WhatsApp", wa, width="stretch")
            b3.link_button("Order by email", mail, width="stretch")

    section("Made by KeyStone")
    st.markdown(
        '<div class="card"><p style="margin:0 0 0.6rem">Every workbook is written by our founder: a '
        "UCT-trained mechanical engineer working at an engineering company, who has spent 5 years preparing "
        "Grade 12 IEB &amp; NSC students in Maths, Physical Sciences and Accounting. The explanations are "
        "the ones that worked in real lessons.</p>"
        '<div class="chips" style="margin:0">'
        '<span class="chip">Mechanical engineer (UCT)</span>'
        '<span class="chip">Working in engineering</span>'
        '<span class="chip">5 years tutoring Grade 12</span>'
        '<span class="chip y">Payments confirmed within 24 hours</span></div></div>',
        unsafe_allow_html=True,
    )
    st.caption(f"Questions before you buy? WhatsApp {PHONE_DISPLAY} or email {EMAIL}.")


def page_checkout():
    products = load("products")
    pid = st.session_state.get("checkout")
    match = products[products["ID"] == pid] if pid else products.iloc[0:0]
    if match.empty:
        st.info("Pick a workbook first.")
        go("Go to the shop", P_SHOP, primary=True, key="co_shop")
        return
    p = match.iloc[0]
    ref = st.session_state.setdefault("order_ref", new_order_ref())

    st.markdown('<div class="hero" style="font-size:2.6rem">Checkout</div>', unsafe_allow_html=True)
    st.markdown(f'<p class="sub"><b>{p["Title"]}</b> · {rand(int(p["Price"]))}</p>', unsafe_allow_html=True)

    wa, mail = order_links(p["Title"], int(p["Price"]))
    st.caption("Prefer to chat first? Order on WhatsApp or by email and we'll help you directly.")
    c1, c2 = st.columns(2)
    c1.link_button("Order on WhatsApp", wa, width="stretch")
    c2.link_button("Order by email", mail, width="stretch")

    section("1. Pay by EFT")
    bank = bank_details()
    if bank:
        rows = "".join(f'<div class="r"><span>{k}</span><b>{v}</b></div>' for k, v in bank.items())
        rows += f'<div class="r"><span>Amount</span><b>{rand(int(p["Price"]))}</b></div>'
        st.markdown(f'<div class="bank">{rows}</div>', unsafe_allow_html=True)
    else:
        st.warning(f"Banking details are being updated. Please WhatsApp us on {PHONE_DISPLAY} to pay.")
    st.markdown(f'Use this as your payment reference: <span class="ref">{ref}</span>', unsafe_allow_html=True)
    st.caption("Instant EFT shows up fastest. Normal EFTs can take 1-2 business days between banks. "
               "We confirm payments within 24 hours of them reflecting.")

    section("2. Upload your proof of payment")
    n = st.session_state.setdefault("co_form_n", 0)
    name = st.text_input("Your name", key=f"co{n}_name")
    c1, c2 = st.columns(2)
    email = c1.text_input("Email", key=f"co{n}_email", help="You'll use this to download your workbook")
    phone = c2.text_input("WhatsApp number", key=f"co{n}_phone")
    pop = st.file_uploader("Proof of payment (PDF or screenshot)", type=UPLOAD_TYPES, key=f"co{n}_pop")
    if st.button("Submit proof of payment", type="primary", key=f"co{n}_go"):
        if not name.strip() or "@" not in email or pop is None:
            st.error("Please add your name, a valid email and your proof of payment.")
        elif pop.size > MAX_UPLOAD_MB * 1024 * 1024:
            st.error(f"That file is over {MAX_UPLOAD_MB} MB. A screenshot works fine.")
        else:
            try:
                with st.spinner("Uploading..."):
                    pop_path = store_file(pop, ref, "pop", folder="payments")
            except Exception:
                st.error(f"Sorry, the upload didn't work. Please try again, or WhatsApp us on {PHONE_DISPLAY}.")
            else:
                append("orders", {
                    "Submitted": datetime.now().strftime("%Y-%m-%d %H:%M"), "Order ref": ref,
                    "Product ID": p["ID"], "Product": p["Title"], "Amount": int(p["Price"]),
                    "Name": name, "Email": email.strip().lower(), "Phone": phone,
                    "Proof of payment": pop_path, "Status": "Pending",
                })
                st.session_state["co_form_n"] = n + 1
                st.session_state["order_done"] = ref
                st.session_state.pop("order_ref", None)
                st.session_state.pop("checkout", None)
                st.switch_page(P_DOWNLOADS)


def page_downloads():
    st.markdown('<div class="hero" style="font-size:2.6rem">My <span class="hl">downloads.</span></div>',
                unsafe_allow_html=True)
    done = st.session_state.pop("order_done", None)
    if done:
        st.success(f"Thank you! We've received your proof of payment for order {done}. "
                   "We'll confirm it (usually within a day) and WhatsApp you when it's ready. "
                   "Then come back here to download.")
    st.markdown('<p class="sub">Enter your order reference and email to get your workbook.</p>',
                unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    ref = c1.text_input("Order reference", value=done or "", placeholder="KS-XXXXX").strip().upper()
    email = c2.text_input("Email").strip().lower()
    if st.button("Find my order", type="primary"):
        _read_raw.clear()   # always check the latest status
        orders = load("orders")
        mine = orders[(orders["Order ref"].str.upper() == ref) & (orders["Email"].str.lower() == email)]
        if mine.empty:
            st.error("We couldn't find that order. Check the reference and email, or WhatsApp us.")
        else:
            o = mine.iloc[-1]
            if o["Status"] == "Approved":
                prod = load("products")
                prod = prod[prod["ID"] == o["Product ID"]]
                if prod.empty or not prod.iloc[0]["File"]:
                    st.error(f"Your payment is confirmed but the file is missing. Please WhatsApp us on {PHONE_DISPLAY}.")
                else:
                    try:
                        data = read_bytes(prod.iloc[0]["File"])
                    except Exception:
                        data = b""
                    if data:
                        try:
                            data = stamp_pdf(data, o["Name"].strip() or o["Email"], o["Order ref"])
                        except Exception:
                            pass   # if stamping ever fails, still give them their workbook
                        st.success(f"Payment confirmed. Enjoy {o['Product']}!")
                        st.download_button("Download your workbook", data,
                                           file_name=f"{_safe(o['Product'])}.pdf", mime="application/pdf",
                                           type="primary")
                    else:
                        st.error(f"Something went wrong fetching the file. Please WhatsApp us on {PHONE_DISPLAY}.")
            elif o["Status"] == "Rejected":
                st.error(f"We couldn't confirm this payment. Please WhatsApp us on {PHONE_DISPLAY}.")
            else:
                st.info("We're still checking your payment. We'll WhatsApp you as soon as it's confirmed.")

# ---------- Leave a review (link you send to parents) ----------
def page_review():
    st.markdown('<div class="hero" style="font-size:2.6rem">How did we <span class="hl">do?</span></div>',
                unsafe_allow_html=True)
    st.markdown('<p class="sub">Your review helps other parents find us. Thank you for taking a minute!</p>',
                unsafe_allow_html=True)
    if st.session_state.pop("review_ok", False):
        st.success("Thank you so much! We really appreciate it.")
        st.balloons()
        return
    n = st.session_state.setdefault("review_form_n", 0)
    name = st.text_input("Your name", key=f"rv{n}_name", help="Only we see this")
    who = st.text_input("How should we show your name?", key=f"rv{n}_who",
                        placeholder="e.g. Thandi M., parent of a Grade 11 learner")
    stars = st.radio("Your rating", [5, 4, 3, 2, 1], horizontal=True, key=f"rv{n}_stars",
                     format_func=lambda x: "★" * x)
    quote = st.text_area("Your review", key=f"rv{n}_quote",
                         placeholder="What changed for your child? What did you like?")
    ok = st.checkbox("KeyStone may show this review on its website and social media", key=f"rv{n}_ok")
    if st.button("Send my review", type="primary", key=f"rv{n}_go"):
        if not name.strip() or len(quote.strip()) < 10:
            st.error("Please add your name and a few words about your experience.")
        elif not ok:
            st.error("Please tick the box so we're allowed to share your review.")
        else:
            append("reviews", {"Quote": quote.strip(), "Who": who.strip() or name.strip().split()[0],
                               "Stars": int(stars), "Show": False,
                               "Submitted": datetime.now().strftime("%Y-%m-%d %H:%M"),
                               "Name (private)": name.strip()})
            st.session_state["review_form_n"] = n + 1
            st.session_state["review_ok"] = True
            st.rerun()

# ---------- Admin ----------
def page_admin():
    if not st.session_state.get("admin"):
        st.subheader("Admin")
        with st.form("login"):
            pw = st.text_input("Password", type="password")
            if st.form_submit_button("Log in"):
                if admin_password() is None:
                    st.error("Admin is locked: no ADMIN_PASSWORD has been set in the app's secrets.")
                elif pw == admin_password():
                    st.session_state["admin"] = True
                    st.rerun()
                else:
                    st.error("Wrong password.")
    else:
        top1, top2 = st.columns([4, 1])
        top1.subheader("Admin")
        if top2.button("Log out"):
            st.session_state["admin"] = False
            st.rerun()

        a_pk, a_bk, a_ap, a_or, a_sh, a_tu, a_rv = st.tabs(
            ["Packages", "Bookings", "Applications", "Orders", "Shop", "Tutors", "Reviews"])

        with a_pk:
            st.write("Track each package here. Tick **Client paid** and **Tutor paid** as money moves.")
            pk = load("packages")
            tutor_names = sorted(load("tutors")["Name"].loc[lambda s: s.str.strip() != ""].unique())
            edited = st.data_editor(
                pk, num_rows="dynamic", width="stretch", hide_index=True, key="pk_edit",
                column_config={
                    "Level": st.column_config.SelectboxColumn(options=LEVELS),
                    "Tutor": st.column_config.SelectboxColumn(options=tutor_names) if tutor_names
                    else st.column_config.TextColumn(),
                    "Subject": st.column_config.SelectboxColumn(options=SUBJECTS + ["Other"]),
                    "Package": st.column_config.SelectboxColumn(options=["Standard"] + list(SPECIALS)),
                    "Package size": st.column_config.NumberColumn(min_value=1, max_value=50, step=1,
                                                                  help="Number of lessons"),
                    "Price (R)": st.column_config.NumberColumn(
                        min_value=0, step=10,
                        help="Leave at 0 for standard packs (worked out from the price list). "
                             "Fill in for special packages."),
                    "Lessons done": st.column_config.NumberColumn(min_value=0, max_value=10, step=1),
                    "Start date": st.column_config.TextColumn(help="e.g. 2026-11-03"),
                },
            )
            if st.button("Save packages", type="primary"):
                save("packages", edited)
                st.success("Saved.")

            view = edited.copy()
            view["Package size"] = pd.to_numeric(view["Package size"], errors="coerce").fillna(0).astype(int)
            view["Lessons done"] = pd.to_numeric(view["Lessons done"], errors="coerce").fillna(0).astype(int)
            view = view[view["Package size"] > 0]
            if not view.empty:
                override = pd.to_numeric(view["Price (R)"], errors="coerce").fillna(0).astype(int)
                view["Value"] = [o if o > 0 else package_value(l, s)
                                 for o, l, s in zip(override, view["Level"], view["Package size"])]
                view["Tutor gets"] = (view["Value"] * TUTOR_SHARE).round().astype(int)
                done = view["Lessons done"] >= view["Package size"]
                active = view[~done]
                to_pay = view[done & ~view["Tutor paid"].astype(bool)]
                m1, m2, m3 = st.columns(3)
                m1.metric("Active packages", len(active))
                m2.metric("Ready to pay tutor", len(to_pay))
                m3.metric("Owed to tutors", rand(int(to_pay["Tutor gets"].sum())))

                if not active.empty:
                    st.markdown("#### In progress")
                    for _, r in active.iterrows():
                        st.progress(
                            min(1.0, r["Lessons done"] / r["Package size"]),
                            text=f"{r['Student'] or 'Unnamed'} with {r['Tutor'] or 'no tutor yet'}: "
                                 f"lesson {r['Lessons done']} of {r['Package size']}",
                        )
                if not to_pay.empty:
                    st.markdown("#### Package complete: pay the tutor")
                    st.dataframe(
                        to_pay[["Student", "Tutor", "Level", "Package size", "Value", "Tutor gets"]],
                        hide_index=True, width="stretch",
                        column_config={"Value": st.column_config.NumberColumn(format="R%d"),
                                       "Tutor gets": st.column_config.NumberColumn(format="R%d")},
                    )

        with a_bk:
            bk = load("bookings")
            st.write(f"**{len(bk)}** booking request(s) from the website, newest first.")
            st.dataframe(bk.iloc[::-1], hide_index=True, width="stretch")
            if not bk.empty:
                st.download_button("Download as CSV", bk.to_csv(index=False), "bookings.csv", "text/csv")

        with a_ap:
            ap = load("applications")
            st.write(f"**{len(ap)}** tutor application(s), newest first.")
            view = ap.iloc[::-1].copy()
            link_cfg = {}
            if USE_SUPABASE:
                for col in ("Matric results", "Uni transcript"):
                    view[col] = [file_link(p) for p in view[col]]
                    link_cfg[col] = st.column_config.LinkColumn(display_text="Open")
                st.caption("Document links work for 1 hour. Reopen this tab for fresh links.")
            elif (DATA / "uploads").exists():
                st.caption(f"Running locally: documents are saved in {DATA / 'uploads'}")
            first = ["Submitted", "Name", "Phone", "Matric results", "Uni transcript"]
            view = view[first + [c for c in view.columns if c not in first]]
            st.dataframe(view, hide_index=True, width="stretch", column_config=link_cfg)
            if not ap.empty:
                st.download_button("Download as CSV", ap.to_csv(index=False), "applications.csv", "text/csv")

        with a_or:
            orders = load("orders")
            pending = int((orders["Status"] == "Pending").sum())
            st.write(f"**{pending}** order(s) waiting for you to check payment.")
            st.caption("Check the money is in your account (match the order reference), then set the "
                       "Status to **Approved** and click Save. The buyer can then download on My downloads.")
            view = orders.iloc[::-1].reset_index(drop=True)
            link_cfg = {"Status": st.column_config.SelectboxColumn(options=["Pending", "Approved", "Rejected"]),
                        "Amount": st.column_config.NumberColumn(format="R%d")}
            if USE_SUPABASE:
                view["Proof link"] = [file_link(x) for x in view["Proof of payment"]]
                link_cfg["Proof link"] = st.column_config.LinkColumn(display_text="Open")
            first = ["Order ref", "Status", "Name", "Amount", "Product"] + (["Proof link"] if USE_SUPABASE else [])
            view = view[first + [c for c in view.columns if c not in first]]
            edited = st.data_editor(view, hide_index=True, width="stretch", key="or_edit",
                                    disabled=[c for c in view.columns if c != "Status"],
                                    column_config=link_cfg)
            if st.button("Save order statuses", type="primary"):
                save("orders", edited.drop(columns=["Proof link"], errors="ignore").iloc[::-1])
                st.success("Saved.")

        with a_sh:
            st.write("Add a workbook: upload the full PDF and the site makes a 3-page preview by itself.")
            with st.form("add_product", clear_on_submit=True):
                title = st.text_input("Title", placeholder="Gr 12 Electrodynamics Workbook")
                c1, c2, c3 = st.columns(3)
                subj = c1.selectbox("Subject", SUBJECTS + ["Other"])
                grade = c2.selectbox("Grade", ["Grade 8", "Grade 9", "Grade 10", "Grade 11", "Grade 12",
                                               "First-year uni"], index=4)
                price = c3.number_input("Price (R)", min_value=0, step=10, value=120)
                desc = st.text_area("Short description (one or two lines)")
                pdf = st.file_uploader("Workbook PDF", type=["pdf"])
                add = st.form_submit_button("Add workbook", type="primary")
            if add:
                if not title.strip() or pdf is None:
                    st.error("Please add a title and the PDF.")
                else:
                    try:
                        with st.spinner("Uploading and making previews..."):
                            pid = _safe(title) + "-" + datetime.now().strftime("%H%M%S")
                            data = pdf.getvalue()
                            file_path = store_bytes(data, f"workbooks/{pid}.pdf", "application/pdf")
                            imgs, n_pages = pdf_previews(data)
                            prev = [store_bytes(img, f"previews/{pid}_{i + 1}.png", "image/png")
                                    for i, img in enumerate(imgs)]
                            append("products", {"ID": pid, "Title": title.strip(), "Subject": subj,
                                                "Grade": grade, "Price": int(price), "Pages": n_pages,
                                                "Description": desc.strip(), "File": file_path,
                                                "Previews": "|".join(prev), "Show": True})
                        st.success(f"Added {title}! It's now in the Shop.")
                    except Exception as e:
                        st.error(f"Couldn't add the workbook: {e}")
            st.markdown("#### Make a stamped copy (for WhatsApp or email orders)")
            st.caption("Once they've paid, pick the workbook, type their name and download a copy with "
                       "'Licensed to ...' on every page, then send it to them.")
            all_prods = load("products")
            all_prods = all_prods[all_prods["Title"].str.strip() != ""]
            if all_prods.empty:
                st.info("Add a workbook first.")
            else:
                m1, m2 = st.columns(2)
                pick = m1.selectbox("Workbook", all_prods["Title"].tolist(), key="manual_pick")
                buyer = m2.text_input("Buyer's name", key="manual_buyer")
                if st.button("Prepare stamped copy", key="manual_make"):
                    if not buyer.strip():
                        st.error("Type the buyer's name first.")
                    else:
                        row = all_prods[all_prods["Title"] == pick].iloc[0]
                        ref = new_order_ref()
                        try:
                            stamped = stamp_pdf(read_bytes(row["File"]), buyer.strip(), ref)
                        except Exception as e:
                            st.error(f"Couldn't prepare the copy: {e}")
                        else:
                            append("orders", {
                                "Submitted": datetime.now().strftime("%Y-%m-%d %H:%M"), "Order ref": ref,
                                "Product ID": row["ID"], "Product": row["Title"], "Amount": int(row["Price"]),
                                "Name": buyer.strip(), "Email": "", "Phone": "",
                                "Proof of payment": "WhatsApp/email order", "Status": "Approved",
                            })
                            st.session_state["manual_copy"] = (stamped, f"{_safe(row['Title'])}_{_safe(buyer)}.pdf", ref)
                if st.session_state.get("manual_copy"):
                    data, fname, ref = st.session_state["manual_copy"]
                    st.success(f"Ready! Order {ref} has been added to your Orders tab.")
                    st.download_button("Download stamped copy", data, file_name=fname,
                                       mime="application/pdf", type="primary", key="manual_dl")

            st.markdown("#### Your workbooks")
            st.caption("Change prices or untick Show to hide one, then click Save.")
            prods = load("products")
            pe = st.data_editor(prods, hide_index=True, width="stretch", key="sh_edit",
                                disabled=["ID", "File", "Previews", "Pages"],
                                column_config={"Price": st.column_config.NumberColumn(format="R%d")})
            if st.button("Save workbooks", type="primary"):
                save("products", pe)
                st.success("Saved.")

        with a_tu:
            st.write("Edit tutor profiles. Untick **Show** to hide someone from the Our tutors page.")
            tu = st.data_editor(load("tutors"), num_rows="dynamic", width="stretch",
                                hide_index=True, key="tu_edit")
            if st.button("Save tutors", type="primary"):
                save("tutors", tu)
                st.success("Saved. Refresh the page to see the changes.")

        with a_rv:
            all_rv = load("reviews")
            waiting = int((~all_rv["Show"] & (all_rv["Quote"].str.strip() != "")).sum())
            st.write(f"**{waiting}** review(s) waiting for you. Tick **Show** on the ones you want on the "
                     "Home page, then click Save.")
            st.caption("Send parents this link to leave a review: your site address + /review "
                       "(e.g. keystonetutoring.streamlit.app/review). You can also add reviews here yourself.")
            rv = st.data_editor(
                all_rv, num_rows="dynamic", width="stretch", hide_index=True,
                key="rv_edit",
                column_config={
                    "Quote": st.column_config.TextColumn(width="large"),
                    "Who": st.column_config.TextColumn(help="e.g. Parent of a Grade 11 learner"),
                    "Stars": st.column_config.NumberColumn(min_value=1, max_value=5, step=1),
                },
            )
            if st.button("Save reviews", type="primary"):
                save("reviews", rv)
                st.success("Saved. Refresh the page to see the changes.")

# ---------- Navigation (menu bar along the top) ----------
P_HOME = st.Page(page_home, title="Home", url_path="home", default=True)
P_TUTORS = st.Page(page_tutors, title="Our tutors", url_path="tutors")
P_PRICING = st.Page(page_pricing, title="Pricing", url_path="pricing")
P_SPECIAL = st.Page(page_special, title="Special packages", url_path="special-packages")
P_QUIZ = st.Page(page_quiz, title="Which package?", url_path="which-package")
P_HOW = st.Page(page_how, title="How it works", url_path="how-it-works")
P_FAQ = st.Page(page_faq, title="FAQ", url_path="faq")
P_BOOK = st.Page(page_book, title="Book a lesson", url_path="book")
P_JOIN = st.Page(page_join, title="Tutor with us", url_path="join")
P_SHOP = st.Page(page_shop, title="Workbooks", url_path="shop")
P_DOWNLOADS = st.Page(page_downloads, title="My downloads", url_path="downloads")
P_CHECKOUT = st.Page(page_checkout, title="Checkout", url_path="checkout", visibility="hidden")
# Admin is hidden from the menu: open it at localhost:8600/admin
P_ADMIN = st.Page(page_admin, title="Admin", url_path="admin", visibility="hidden")
P_REVIEW = st.Page(page_review, title="Leave a review", url_path="review", visibility="hidden")

nav = st.navigation(
    {
        "": [P_HOME, P_ADMIN, P_CHECKOUT, P_REVIEW],
        "About us": [P_TUTORS, P_HOW, P_FAQ],
        "Packages & pricing": [P_PRICING, P_SPECIAL, P_QUIZ],
        "Shop": [P_SHOP, P_DOWNLOADS],
        "Get started": [P_BOOK, P_JOIN],
    },
    position="top",
)
nav.run()

# ---------- Footer ----------
st.markdown(
    f'<div class="foot">@{INSTAGRAM} &nbsp;·&nbsp; {PHONE_DISPLAY} &nbsp;·&nbsp; {EMAIL}</div>',
    unsafe_allow_html=True,
)
_, mid, _ = st.columns([2, 1, 2])
mid.page_link(P_ADMIN, label="Admin login", icon=":material/lock:")
