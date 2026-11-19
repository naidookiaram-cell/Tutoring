"""KeyStone Tutors - Streamlit site.

Run it with:   streamlit run app.py   (or double-click "Start KeyStone Site")

Everything the site saves (bookings, tutor applications, packages, reviews,
tutor profiles) goes to your Google Sheet when SHEET_URL and
[gcp_service_account] are set in the app's secrets. Without them (e.g. on
your laptop) it falls back to CSV files in the "data" folder.
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

# Special packages. Set "price" to a number (e.g. 3200) to show it on the site;
# None shows "Ask us for a quote". Add, remove or rename freely.
SPECIALS = {
    "Exam Prep": {
        "for": "Grades 10-12 & first-year uni",
        "blurb": "A focused run-up to exams, so nothing on the paper is a surprise.",
        "includes": ["Past papers worked through together", "Extra time on the topics they dread",
                     "Exam technique and time management"],
        "price": None,
    },
    "NBT Prep": {
        "for": "Matrics applying to university",
        "blurb": "Get ready for the National Benchmark Tests many universities ask for.",
        "includes": ["Academic & Quantitative Literacy (AQL)", "Mathematics (MAT)",
                     "Timed practice under test conditions"],
        "price": None,
    },
    "Holiday Catch-up": {
        "for": "Grades 1-12",
        "blurb": "Use the school holidays to close last term's gaps or get ahead of next term.",
        "includes": ["Recap of topics that didn't stick", "A preview of next term's work",
                     "Holiday-friendly session times"],
        "price": None,
    },
    "Varsity Head Start": {
        "for": "Matrics heading to university",
        "blurb": "Bridge the jump from school to first-year maths, science or accounting.",
        "includes": ["First-year topics before lectures start", "Study habits for a uni workload",
                     "Tips from tutors who've just done first year"],
        "price": None,
    },
}

# Tutor cards only appear once at least this many tutors are ticked "Show".
MIN_TUTORS_TO_SHOW = 3

# ---------- Data files ----------
DATA = Path(__file__).parent / "data"
DATA.mkdir(exist_ok=True)

SCHEMAS = {
    "bookings": {"Submitted": "text", "Parent": "text", "Student": "text", "Level": "text",
                 "Subject": "text", "Package": "text", "Notes": "text"},
    "applications": {"Submitted": "text", "Name": "text", "Email": "text", "Phone": "text",
                     "University": "text", "Degree and year": "text", "Subjects": "text",
                     "Levels": "text", "Marks": "text", "Availability": "text", "Why KeyStone": "text"},
    "packages": {"Start date": "text", "Student": "text", "Parent": "text", "Tutor": "text",
                  "Subject": "text", "Level": "text", "Package": "text", "Package size": "int",
                 "Price (R)": "int", "Lessons done": "int",
                 "Client paid": "bool", "Tutor paid": "bool"},
    "reviews": {"Quote": "text", "Who": "text", "Stars": "int", "Show": "bool"},
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


def admin_password() -> str | None:
    """The admin password comes ONLY from Streamlit secrets. No secret = admin stays locked."""
    try:
        pw = str(st.secrets["ADMIN_PASSWORD"]).strip()
        return pw or None
    except Exception:
        return None


def rand(n: int) -> str:
    return f"R{n:,}".replace(",", " ")


def package_value(level: str, size: int) -> int:
    single, four, ten = PRICES.get(level, (0, 0, 0))
    per = ten if size >= 10 else four if size >= 4 else single
    return per * size


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
      .foot { font-family: 'Space Mono', monospace; font-weight: 700; font-size: 0.8rem; letter-spacing: 0.08em; color: #6B635A; text-align: center; margin-top: 3rem; }
      @media (max-width: 640px) { .hero { font-size: 2.6rem; } }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="kicker">KEYSTONE TUTORS</div>', unsafe_allow_html=True)


# ---------- Home ----------
def page_home():
    left, right = st.columns([3, 2], vertical_alignment="center")
    with left:
        st.markdown(
            '<div class="hero">Tutoring that actually<br><span class="hl">clicks.</span></div>'
            '<div class="hand">run by students, for students</div>',
            unsafe_allow_html=True,
        )
    with right:
        st.markdown(MASCOT, unsafe_allow_html=True)

    st.write(
        "Private one-on-one tutoring, online, for **Grades 1 to 12 and first-year uni**. "
        "IEB and NSC. Our tutors are students who were in your shoes not so long ago, "
        "so they explain things the way they wish someone had explained them."
    )
    c1, c2, c3 = st.columns(3)
    for col, title, text in [
        (c1, "One-on-one", "Every session is just your child and their tutor. No crowded classroom."),
        (c2, "Online", "Learn from home, no driving across town in traffic."),
        (c3, "Matched tutors", "We pair each student with a tutor who suits their subject and grade."),
    ]:
        col.markdown(f'<div class="card"><h4>{title}</h4>{text}</div>', unsafe_allow_html=True)

    st.subheader("Subjects we cover")
    gcols = st.columns(len(SUBJECT_GROUPS))
    for col, (group, subs) in zip(gcols, SUBJECT_GROUPS.items()):
        col.markdown(f'<div class="card"><div class="meta">{group}</div>'
                     + "".join(f"<div>{s}</div>" for s in subs) + "</div>", unsafe_allow_html=True)

    st.markdown(
        '<div class="result"><h4 style="margin:0 0 0.3rem;font-family:Fraunces,Georgia,serif">'
        "Exams coming up? NBTs on the horizon?</h4>"
        "Have a look at our <b>Special packages</b>: Exam Prep, NBT Prep, Holiday Catch-up "
        "and Varsity Head Start.</div>",
        unsafe_allow_html=True,
    )
    st.page_link(P_SPECIAL, label="See special packages", icon=":material/arrow_forward:")

    st.markdown(
        '<div class="card" style="margin-top:1rem"><div class="kicker" style="color:#D9705A">WE\'RE HIRING</div>'
        '<h4>Are you a uni student who\'s good at explaining things?</h4>'
        "Tutor with KeyStone, choose your own hours and earn while you study.</div>",
        unsafe_allow_html=True,
    )
    st.page_link(P_JOIN, label="Find out more and apply", icon=":material/arrow_forward:")

    reviews = load("reviews")
    reviews = reviews[reviews["Show"] & (reviews["Quote"].str.strip() != "")]
    if not reviews.empty:
        st.subheader("What parents say")
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
    st.subheader("Our tutors")
    st.write(
        "KeyStone tutors are university students who were sitting the same exams not long ago, "
        "so they remember exactly what made a topic finally click."
    )
    st.markdown("#### What we look for")
    w1, w2 = st.columns(2)
    for col, title, text in [
        (w1, "Strong in their subjects", "Tutors only teach subjects they've done well in themselves."),
        (w2, "Great at explaining", "Knowing the work isn't enough. They have to make it make sense."),
        (w1, "Matched to your child", "We pair each student with a tutor who suits their subject, grade and style."),
        (w2, "Managed by KeyStone", "We stay in the loop on every student, so you always have someone to talk to."),
    ]:
        col.markdown(f'<div class="card"><h4>{title}</h4>{text}</div>', unsafe_allow_html=True)

    tutors = load("tutors")
    tutors = tutors[tutors["Show"] & (tutors["Name"].str.strip() != "")]
    if len(tutors) < MIN_TUTORS_TO_SHOW:
        tutors = tutors.iloc[0:0]
    else:
        st.markdown("#### Meet the team")
    cols = st.columns(2)
    for i, (_, t) in enumerate(tutors.iterrows()):
        study = " · ".join(x for x in [t["Degree"], t["University"]] if x.strip())
        lines = [f"<h4>{t['Name']}</h4>"]
        if study:
            lines.append(f'<div class="meta">{study}</div>')
        if t["Subjects"].strip():
            lines.append(f"<p><b>Teaches:</b> {t['Subjects']}</p>")
        if t["Grades"].strip():
            lines.append(f"<p><b>Grades:</b> {t['Grades']}</p>")
        if t["Fun fact"].strip():
            lines.append(f'<div class="hand" style="font-size:1.5rem">{t["Fun fact"]}</div>')
        cols[i % 2].markdown(f'<div class="card">{"".join(lines)}</div>', unsafe_allow_html=True)

# ---------- Pricing ----------
def page_pricing():
    st.subheader("Pricing")
    st.write("Prices are **per session**. The bigger the pack, the less each lesson costs.")
    st.table(pd.DataFrame(
        {"Single": [rand(p[0]) for p in PRICES.values()],
         "4-pack": [rand(p[1]) for p in PRICES.values()],
         "10-pack": [rand(p[2]) for p in PRICES.values()]},
        index=pd.Index(LEVELS, name="Level"),
    ))

    st.markdown("#### Quick quote")
    q1, q2 = st.columns(2)
    level = q1.selectbox("Level", LEVELS, key="q_level")
    pack = q2.selectbox("Package", list(PACK_SIZES), key="q_pack")
    size = PACK_SIZES[pack]
    total = package_value(level, size)
    st.metric("Package total", rand(total), f"{rand(total // size)} per session")
    st.page_link(P_SPECIAL, label="Looking for exam or NBT prep? See special packages",
                 icon=":material/arrow_forward:")

# ---------- Special packages ----------
def page_special():
    st.subheader("Special packages")
    st.write("For the moments that matter most. Each one is planned around your child and their goal.")
    scols = st.columns(2)
    for i, (name, sp) in enumerate(SPECIALS.items()):
        price = rand(sp["price"]) if sp.get("price") else "Ask us for a quote"
        items = "".join(f"<li>{x}</li>" for x in sp["includes"])
        scols[i % 2].markdown(
            f'<div class="card"><div class="meta">{sp["for"]}</div><h4>{name}</h4>'
            f'<p>{sp["blurb"]}</p><ul style="margin:0 0 0.6rem 1.1rem;padding:0">{items}</ul>'
            f'<span class="hl" style="font-weight:700">{price}</span></div>',
            unsafe_allow_html=True,
        )
    st.page_link(P_BOOK, label="Interested? Book a lesson", icon=":material/arrow_forward:")

# ---------- Which package? ----------
def page_quiz():
    st.subheader("Which package suits us?")
    st.write("Answer three quick questions and we'll point you in the right direction.")
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
        sp_price = SPECIALS[rec].get("price")
        cost = f"<b>{rand(sp_price)}</b>" if sp_price else "Ask us for a quote. Pricing depends on how much prep is needed."
    else:
        size = PACK_SIZES[rec]
        total = package_value(qa, size)
        cost = f"<b>{rand(total)}</b> in total ({rand(total // size)} per session for {qa})."
    st.markdown(
        f'<div class="result"><div class="meta" style="color:#2A2622">We recommend</div>'
        f'<h3 style="margin:0.2rem 0">{rec}</h3><p>{why}</p>{cost}</div>',
        unsafe_allow_html=True,
    )
    st.page_link(P_BOOK, label="Ready? Book a lesson", icon=":material/arrow_forward:")

# ---------- How it works ----------
def page_how():
    st.subheader("How it works")
    for num, title, text in [
        ("01", "Get in touch", "Send us a message on the Book a lesson page, or WhatsApp us directly."),
        ("02", "Meet your tutor", "We match your child with a tutor and set up a group chat with you, the tutor and KeyStone."),
        ("03", "Choose a package", "Pick a single lesson, a 4-pack or a 10-pack, whatever suits you."),
        ("04", "Start learning", "Your tutor shares session times in the group, and you're off."),
    ]:
        st.markdown(f'<div class="card"><span class="step">{num}</span><h4>{title}</h4>{text}</div>',
                    unsafe_allow_html=True)

# ---------- FAQ ----------
def page_faq():
    st.subheader("Questions parents ask")
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
    st.subheader("Book a lesson")
    st.write("Fill this in and we'll open WhatsApp with your message ready to send.")
    with st.form("booking"):
        parent = st.text_input("Parent / guardian name")
        student = st.text_input("Student name")
        b1, b2 = st.columns(2)
        grade = b1.selectbox("Level", LEVELS)
        subject = b2.selectbox("Subject", SUBJECT_CHOICES)
        package = st.selectbox("Package", list(PACK_SIZES) + list(SPECIALS) + ["Not sure yet"])
        notes = st.text_area("Anything we should know? (topics, exam dates, availability)")
        submitted = st.form_submit_button("Create my message")

    if submitted:
        if not parent.strip() or not student.strip():
            st.error("Please add the parent and student names.")
        else:
            append("bookings", {
                "Submitted": datetime.now().strftime("%Y-%m-%d %H:%M"), "Parent": parent,
                "Student": student, "Level": grade, "Subject": subject,
                "Package": package, "Notes": notes,
            })
            msg = (
                f"Hi KeyStone! I'm {parent} and I'd like to book tutoring for {student}.\n"
                f"Level: {grade}\nSubject: {subject}\nPackage: {package}"
                + (f"\nNotes: {notes}" if notes else "")
            )
            st.success("Your message is ready!")
            st.code(msg, language=None)
            st.link_button("Send on WhatsApp", f"https://wa.me/{WHATSAPP}?text={quote(msg)}", type="primary")
            st.caption(f"Prefer email? Send it to {EMAIL}.")

# ---------- Join our team ----------
def page_join():
    st.markdown(
        '<div class="kicker" style="color:#D9705A">WE\'RE HIRING TUTORS</div>'
        '<div class="hero" style="font-size:3rem">Earn while you study.<br>'
        '<span class="hl">Teach what you\'re good at.</span></div>',
        unsafe_allow_html=True,
    )
    st.write(
        "KeyStone was started by a student, for students. If you're at university and you're "
        "good at explaining things, we'd love to have you on the team."
    )

    st.markdown("#### Why tutor with KeyStone")
    w1, w2 = st.columns(2)
    for col, title, text in [
        (w1, "You keep the bigger share", "You earn the majority of every session fee, not a small cut."),
        (w2, "We find the students", "No advertising or chasing parents. We match students to you."),
        (w1, "Work around your timetable", "You choose how many students you take and when you teach."),
        (w2, "Online, from anywhere", "Teach from res, home or campus. All you need is a laptop and Wi-Fi."),
        (w1, "We handle the admin", "Bookings, payments and parent questions go through KeyStone."),
        (w2, "Great for your CV", "Real teaching and communication experience employers notice."),
    ]:
        col.markdown(f'<div class="card"><h4>{title}</h4>{text}</div>', unsafe_allow_html=True)

    st.markdown("#### Who we're looking for")
    st.markdown(
        '<div class="card">'
        "<p>✓ &nbsp;You're currently studying at a university or college</p>"
        "<p>✓ &nbsp;You did well in the subjects you want to teach</p>"
        "<p>✓ &nbsp;You're patient and can explain things in more than one way</p>"
        "<p style='margin:0'>✓ &nbsp;You're reliable: you show up on time and reply to messages</p>"
        "</div>",
        unsafe_allow_html=True,
    )

    st.markdown("#### How it works")
    steps = [
        ("01", "Apply", "Fill in the form below. It takes about 5 minutes."),
        ("02", "Chat with us", "We'll get in touch to get to know you and talk through how it works."),
        ("03", "Get matched", "We pair you with students in your subjects and grades."),
        ("04", "Teach & track", "Run your sessions online and fill in a simple timesheet per package."),
        ("05", "Get paid", "You're paid once each package of lessons is complete."),
    ]
    scols = st.columns(len(steps))
    for col, (num, title, text) in zip(scols, steps):
        col.markdown(f'<div class="card"><span class="step">{num}</span><h4 style="font-size:1.05rem">'
                     f"{title}</h4><small>{text}</small></div>", unsafe_allow_html=True)

    st.markdown("#### Apply now")
    with st.form("apply", clear_on_submit=True):
        a1, a2 = st.columns(2)
        name = a1.text_input("Full name")
        phone = a2.text_input("WhatsApp number")
        email = st.text_input("Email")
        a3, a4 = st.columns(2)
        uni = a3.text_input("University")
        degree = a4.text_input("Degree and year (e.g. BSc Eng, 3rd year)")
        subj = st.multiselect("Subjects you can tutor", SUBJECTS + ["Other"])
        lvls = st.multiselect("Levels you can tutor", LEVELS)
        marks = st.text_area("Your marks in those subjects (matric or uni)")
        avail = st.text_input("When are you usually free? (e.g. weekday evenings, Saturdays)")
        why = st.text_area("Why would you like to tutor with KeyStone?")
        applied = st.form_submit_button("Send my application")

    if applied:
        if not name.strip() or not (phone.strip() or email.strip()):
            st.error("Please add your name and a way to contact you.")
        else:
            append("applications", {
                "Submitted": datetime.now().strftime("%Y-%m-%d %H:%M"), "Name": name, "Email": email,
                "Phone": phone, "University": uni, "Degree and year": degree,
                "Subjects": ", ".join(subj), "Levels": ", ".join(lvls), "Marks": marks,
                "Availability": avail, "Why KeyStone": why,
            })
            st.success("Thanks for applying! We'll be in touch on WhatsApp or email.")

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

        a_pk, a_bk, a_ap, a_tu, a_rv = st.tabs(
            ["Packages", "Bookings", "Applications", "Tutors", "Reviews"])

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
            st.dataframe(ap.iloc[::-1], hide_index=True, width="stretch")
            if not ap.empty:
                st.download_button("Download as CSV", ap.to_csv(index=False), "applications.csv", "text/csv")

        with a_tu:
            st.write("Edit tutor profiles. Untick **Show** to hide someone from the Our tutors page.")
            tu = st.data_editor(load("tutors"), num_rows="dynamic", width="stretch",
                                hide_index=True, key="tu_edit")
            if st.button("Save tutors", type="primary"):
                save("tutors", tu)
                st.success("Saved. Refresh the page to see the changes.")

        with a_rv:
            st.write("Paste in real messages from happy parents. Ticked ones show on the Home page.")
            rv = st.data_editor(
                load("reviews"), num_rows="dynamic", width="stretch", hide_index=True,
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
# Admin is hidden from the menu: open it at localhost:8600/admin
P_ADMIN = st.Page(page_admin, title="Admin", url_path="admin", visibility="hidden")

nav = st.navigation(
    {
        "": [P_HOME, P_ADMIN],
        "About us": [P_TUTORS, P_HOW, P_FAQ],
        "Packages & pricing": [P_PRICING, P_SPECIAL, P_QUIZ],
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
