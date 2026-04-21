import streamlit as st
import pandas as pd
import uuid
import time
import io
import json
import hashlib
import datetime
import re
from docx import Document

# ─────────────────────────────────────────────
# PAGE CONFIG  (must be first Streamlit call)
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="QA Forge – AI Test Case Engine",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────
# GLOBAL CSS  –  cyberpunk / dark-forge theme
# ─────────────────────────────────────────────
GLOBAL_CSS = """
<link href="https://fonts.googleapis.com/css2?family=Orbitron:wght@400;700;900&family=Space+Mono:ital,wght@0,400;0,700;1,400&family=Inter:wght@300;400;500;600&display=swap" rel="stylesheet">

<style>
/* ── reset & root ─────────────────────────── */
:root {
    --bg:        #050810;
    --surface:   #0c1222;
    --panel:     #111827;
    --border:    #1e2d47;
    --accent:    #00d4ff;
    --accent2:   #7c3aed;
    --accent3:   #f59e0b;
    --danger:    #ef4444;
    --success:   #10b981;
    --text:      #e2e8f0;
    --muted:     #64748b;
    --glow:      0 0 20px rgba(0,212,255,.35);
    --glow2:     0 0 20px rgba(124,58,237,.35);
}

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
    background: var(--bg) !important;
    color: var(--text) !important;
}

/* hide streamlit chrome */
#MainMenu, footer, header { visibility: hidden; }
.block-container { padding: 0 !important; max-width: 100% !important; }

/* ── particle canvas bg ───────────────────── */
.qaforge-bg {
    position: fixed; inset: 0; z-index: 0;
    background: radial-gradient(ellipse 80% 60% at 20% 20%, rgba(0,212,255,.08) 0%, transparent 55%),
                radial-gradient(ellipse 60% 50% at 80% 80%, rgba(124,58,237,.10) 0%, transparent 55%),
                radial-gradient(ellipse 70% 70% at 50% 50%, rgba(5,8,16,1) 0%, #050810 100%);
}

/* ── LOGIN PAGE ──────────────────────────── */
.login-shell {
    position: relative; z-index: 10;
    display: flex; align-items: center; justify-content: center;
    min-height: 100vh; width: 100%;
    padding: 2rem;
}

.login-card {
    background: linear-gradient(135deg, rgba(12,18,34,.95) 0%, rgba(17,24,39,.95) 100%);
    border: 1px solid var(--border);
    border-radius: 20px;
    padding: 3rem 2.8rem 2.5rem;
    width: 100%; max-width: 440px;
    box-shadow: 0 0 0 1px rgba(0,212,255,.06),
                0 40px 80px rgba(0,0,0,.6),
                inset 0 1px 0 rgba(255,255,255,.04);
    position: relative; overflow: hidden;
}

.login-card::before {
    content: '';
    position: absolute; top: 0; left: 0; right: 0; height: 2px;
    background: linear-gradient(90deg, transparent, var(--accent), var(--accent2), transparent);
    animation: shimmer 3s linear infinite;
}
@keyframes shimmer { 0%{opacity:.4} 50%{opacity:1} 100%{opacity:.4} }

.login-logo {
    text-align: center; margin-bottom: 2.2rem;
}
.login-logo .logo-icon {
    font-size: 3.2rem; display: block;
    filter: drop-shadow(0 0 16px rgba(0,212,255,.8));
    animation: pulse-glow 2.5s ease-in-out infinite;
}
@keyframes pulse-glow {
    0%,100%{ filter: drop-shadow(0 0 16px rgba(0,212,255,.6)); }
    50%{     filter: drop-shadow(0 0 28px rgba(0,212,255,1)); }
}
.login-logo h1 {
    font-family: 'Orbitron', monospace;
    font-size: 1.9rem; font-weight: 900;
    background: linear-gradient(90deg, var(--accent), var(--accent2));
    -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    margin: .4rem 0 .15rem; letter-spacing: 2px;
}
.login-logo p {
    font-size: .78rem; color: var(--muted);
    font-family: 'Space Mono', monospace; letter-spacing: 1px;
}

.login-field { margin-bottom: 1.1rem; }
.login-field label {
    font-size: .75rem; font-weight: 600;
    color: var(--accent); letter-spacing: 1.5px;
    text-transform: uppercase; display: block; margin-bottom: .4rem;
    font-family: 'Space Mono', monospace;
}

/* style streamlit text_input inside login */
.login-card .stTextInput > div > div > input {
    background: rgba(0,212,255,.04) !important;
    border: 1px solid var(--border) !important;
    border-radius: 10px !important;
    color: var(--text) !important;
    font-family: 'Space Mono', monospace !important;
    font-size: .88rem !important;
    padding: .65rem 1rem !important;
    transition: border-color .25s, box-shadow .25s;
}
.login-card .stTextInput > div > div > input:focus {
    border-color: var(--accent) !important;
    box-shadow: var(--glow) !important;
    outline: none !important;
}

/* login button */
.login-card .stButton > button {
    width: 100% !important;
    background: linear-gradient(135deg, var(--accent2) 0%, var(--accent) 100%) !important;
    color: #fff !important;
    border: none !important; border-radius: 10px !important;
    font-family: 'Orbitron', monospace !important;
    font-weight: 700 !important; letter-spacing: 2px !important;
    font-size: .85rem !important; padding: .75rem !important;
    cursor: pointer !important;
    transition: transform .15s, box-shadow .25s !important;
}
.login-card .stButton > button:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 0 30px rgba(0,212,255,.5) !important;
}

.login-error {
    background: rgba(239,68,68,.1);
    border: 1px solid rgba(239,68,68,.3);
    border-radius: 8px; padding: .6rem 1rem;
    font-size: .82rem; color: var(--danger);
    margin-top: .8rem; text-align: center;
    font-family: 'Space Mono', monospace;
}

/* demo hint */
.demo-hint {
    text-align: center; margin-top: 1.5rem; padding-top: 1.2rem;
    border-top: 1px solid var(--border);
    font-size: .72rem; color: var(--muted);
    font-family: 'Space Mono', monospace;
}
.demo-hint span { color: var(--accent); }

/* ── APP SHELL ────────────────────────────── */
.app-shell { position: relative; z-index: 10; padding: 0 2rem 3rem; }

/* top navbar */
.top-nav {
    display: flex; align-items: center; justify-content: space-between;
    padding: 1rem 2rem;
    background: rgba(12,18,34,.9);
    border-bottom: 1px solid var(--border);
    backdrop-filter: blur(12px);
    position: sticky; top: 0; z-index: 999;
}
.nav-logo {
    font-family: 'Orbitron', monospace; font-weight: 900; font-size: 1.2rem;
    background: linear-gradient(90deg, var(--accent), var(--accent2));
    -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    letter-spacing: 2px;
}
.nav-right { display: flex; align-items: center; gap: 1.2rem; }
.nav-badge {
    font-size: .7rem; font-family: 'Space Mono', monospace;
    background: rgba(0,212,255,.1); border: 1px solid rgba(0,212,255,.25);
    color: var(--accent); padding: .2rem .65rem; border-radius: 20px;
}
.nav-user {
    font-size: .8rem; color: var(--muted);
    font-family: 'Space Mono', monospace;
}

/* ── STAT CARDS ───────────────────────────── */
.stat-row { display: grid; grid-template-columns: repeat(4,1fr); gap: 1rem; margin: 1.5rem 0; }
.stat-card {
    background: var(--panel);
    border: 1px solid var(--border);
    border-radius: 14px; padding: 1.2rem 1.4rem;
    position: relative; overflow: hidden;
    transition: transform .2s, border-color .25s;
}
.stat-card:hover { transform: translateY(-3px); border-color: var(--accent); }
.stat-card::after {
    content: ''; position: absolute;
    inset: 0; border-radius: 14px;
    background: linear-gradient(135deg, rgba(0,212,255,.04), transparent);
    pointer-events: none;
}
.stat-num {
    font-family: 'Orbitron', monospace; font-size: 2rem; font-weight: 900;
    line-height: 1;
    background: linear-gradient(90deg, var(--accent), var(--accent2));
    -webkit-background-clip: text; -webkit-text-fill-color: transparent;
}
.stat-label {
    font-size: .72rem; color: var(--muted); letter-spacing: 1.2px;
    text-transform: uppercase; margin-top: .35rem;
    font-family: 'Space Mono', monospace;
}

/* ── UPLOAD ZONE ──────────────────────────── */
.upload-zone {
    background: var(--panel);
    border: 2px dashed var(--border);
    border-radius: 16px; padding: 2.5rem;
    text-align: center; transition: border-color .3s;
}
.upload-zone:hover { border-color: var(--accent); }

/* ── SECTION HEADER ───────────────────────── */
.section-header {
    font-family: 'Orbitron', monospace; font-weight: 700; font-size: 1.05rem;
    color: var(--text); letter-spacing: 1.5px;
    display: flex; align-items: center; gap: .6rem;
    margin-bottom: 1rem; padding-bottom: .7rem;
    border-bottom: 1px solid var(--border);
}
.section-header .dot {
    width: 8px; height: 8px; border-radius: 50%;
    background: var(--accent);
    box-shadow: 0 0 8px var(--accent);
    flex-shrink: 0;
}

/* ── FEATURE PILLS ────────────────────────── */
.feature-pills { display: flex; flex-wrap: wrap; gap: .5rem; margin-bottom: 1.2rem; }
.fpill {
    font-size: .72rem; font-family: 'Space Mono', monospace;
    padding: .25rem .7rem; border-radius: 20px; border: 1px solid;
    white-space: nowrap;
}
.fpill-blue  { border-color:rgba(0,212,255,.3);  color:var(--accent);  background:rgba(0,212,255,.07); }
.fpill-purple{ border-color:rgba(124,58,237,.3); color:#a78bfa;        background:rgba(124,58,237,.07); }
.fpill-amber { border-color:rgba(245,158,11,.3); color:var(--accent3); background:rgba(245,158,11,.07); }
.fpill-green { border-color:rgba(16,185,129,.3); color:var(--success); background:rgba(16,185,129,.07); }

/* ── PROGRESS BAR ─────────────────────────── */
.prog-wrap { background: rgba(255,255,255,.05); border-radius: 99px; height:6px; overflow:hidden; margin:.6rem 0; }
.prog-fill  { height:100%; border-radius:99px; background:linear-gradient(90deg,var(--accent2),var(--accent)); transition:width .4s; }

/* ── GENERATE BUTTON override ─────────────── */
.stButton > button[kind="primary"],
div[data-testid="stButton"] > button {
    background: linear-gradient(135deg, var(--accent2) 0%, var(--accent) 100%) !important;
    color: #fff !important; border: none !important;
    border-radius: 10px !important;
    font-family: 'Orbitron', monospace !important;
    font-weight: 700 !important; letter-spacing: 1.5px !important;
    padding: .65rem 1.8rem !important;
    transition: transform .15s, box-shadow .25s !important;
}
.stButton > button:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 0 24px rgba(0,212,255,.45) !important;
}

/* ── DATAFRAME ────────────────────────────── */
.dataframe-container .stDataFrame {
    border: 1px solid var(--border) !important;
    border-radius: 12px !important;
}

/* ── SIDEBAR ──────────────────────────────── */
[data-testid="stSidebar"] {
    background: var(--surface) !important;
    border-right: 1px solid var(--border) !important;
}

/* ── SELECTBOX / TABS override ─────────────── */
.stTabs [data-baseweb="tab-list"] {
    background: var(--panel) !important;
    border-radius: 10px !important; gap: 4px; padding: 4px;
}
.stTabs [data-baseweb="tab"] {
    font-family: 'Space Mono', monospace !important;
    font-size: .78rem !important; color: var(--muted) !important;
    background: transparent !important;
    border-radius: 8px !important; padding: .4rem 1rem !important;
}
.stTabs [aria-selected="true"] {
    background: linear-gradient(135deg,var(--accent2),var(--accent)) !important;
    color: #fff !important;
}

/* ── SUCCESS / INFO banners ───────────────── */
.banner {
    border-radius: 12px; padding: .85rem 1.2rem;
    display: flex; align-items: center; gap: .75rem;
    font-size: .84rem; font-family: 'Space Mono', monospace;
    margin: .8rem 0;
}
.banner-success { background:rgba(16,185,129,.1); border:1px solid rgba(16,185,129,.3); color:var(--success); }
.banner-info    { background:rgba(0,212,255,.07);  border:1px solid rgba(0,212,255,.2);  color:var(--accent); }
.banner-warn    { background:rgba(245,158,11,.08); border:1px solid rgba(245,158,11,.25); color:var(--accent3); }

/* scrollbar */
::-webkit-scrollbar { width:6px; height:6px; }
::-webkit-scrollbar-track { background:var(--surface); }
::-webkit-scrollbar-thumb { background:var(--border); border-radius:99px; }
::-webkit-scrollbar-thumb:hover { background:var(--muted); }

/* ── history card ─────────────────────────── */
.hist-card {
    background: var(--panel); border: 1px solid var(--border);
    border-radius: 12px; padding: 1rem 1.2rem;
    margin-bottom: .7rem; cursor: pointer;
    transition: border-color .2s, transform .15s;
}
.hist-card:hover { border-color: var(--accent); transform: translateX(4px); }
.hist-meta { font-size:.7rem; color:var(--muted); font-family:'Space Mono',monospace; }
.hist-title { font-size:.88rem; font-weight:600; color:var(--text); margin-top:.2rem; }

/* ── analytics grid ───────────────────────── */
.analytics-grid { display:grid; grid-template-columns:1fr 1fr; gap:1rem; }
.chart-card {
    background:var(--panel); border:1px solid var(--border);
    border-radius:14px; padding:1.2rem;
}
.chart-card h4 {
    font-family:'Orbitron',monospace; font-size:.8rem; color:var(--muted);
    letter-spacing:1.5px; text-transform:uppercase; margin-bottom:1rem;
}

/* inline bar chart */
.bar-row { display:flex; align-items:center; gap:.8rem; margin:.4rem 0; }
.bar-label { font-size:.75rem; color:var(--muted); width:100px; font-family:'Space Mono',monospace; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
.bar-track { flex:1; background:rgba(255,255,255,.05); border-radius:99px; height:8px; }
.bar-fill  { height:100%; border-radius:99px; }

/* ── log terminal ─────────────────────────── */
.terminal {
    background:#030609; border:1px solid var(--border);
    border-radius:12px; padding:1rem 1.2rem;
    font-family:'Space Mono',monospace; font-size:.75rem;
    color:#4ade80; height:180px; overflow-y:auto;
    line-height:1.7;
}
.terminal .t-dim  { color:var(--muted); }
.terminal .t-blue { color:var(--accent); }
.terminal .t-warn { color:var(--accent3); }
</style>
"""

# ─────────────────────────────────────────────
# USERS DB  (hashed passwords)
# ─────────────────────────────────────────────
def _hash(pw: str) -> str:
    return hashlib.sha256(pw.encode()).hexdigest()

USERS = {
    "admin":   {"password": _hash("admin123"), "role": "Admin",    "name": "Admin User"},
    "qaforge": {"password": _hash("forge2024"), "role": "Senior QA","name": "QA Lead"},
    "demo":    {"password": _hash("demo"),      "role": "Viewer",   "name": "Demo Account"},
}

# ─────────────────────────────────────────────
# SESSION DEFAULTS
# ─────────────────────────────────────────────
for key, val in {
    "authenticated": False,
    "username": "",
    "history": [],          # list of {id, name, ts, df, config}
    "active_df": None,
    "total_generated": 0,
    "login_attempts": 0,
}.items():
    if key not in st.session_state:
        st.session_state[key] = val

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
COLUMNS = [
    "CT Test Case IDs","OSD Number","ACC FLOW ID","ACC",
    "Test Case Description","Test case steps","Expected results",
    "Pre-condition","Post condition","Dev Status","Testing Status",
    "Comments","ETA of Dev","ETA of CT","TESTER"
]

def read_uploaded(file) -> str:
    if file.name.endswith(".docx"):
        doc = Document(file)
        return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    return file.read().decode("utf-8")

def make_ct_id(prefix: str, n: int) -> str:
    return f"{prefix}-{str(n).zfill(4)}"

def parse_pipe_table(raw: str, prefix: str) -> pd.DataFrame:
    rows = []
    ct_counter = 1
    for line in raw.split("\n"):
        if "|" not in line:
            continue
        cols = [c.strip() for c in line.split("|")]
        cols = [c for c in cols if c != ""]   # remove empty edge-split
        if len(cols) >= 14:
            cols = cols[:14]
            ct_id = make_ct_id(prefix, ct_counter)
            rows.append([ct_id] + cols[1:] if cols[0].upper().startswith("CT") else [ct_id] + cols[:14])
            ct_counter += 1
    return pd.DataFrame(rows, columns=COLUMNS) if rows else pd.DataFrame(columns=COLUMNS)

def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Test Cases")
        ws = writer.sheets["Test Cases"]
        for col_cells in ws.columns:
            max_len = max(len(str(c.value or "")) for c in col_cells)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 4, 40)
    return buf.getvalue()

def log_line(msg: str, level: str = "info") -> str:
    ts = datetime.datetime.now().strftime("%H:%M:%S")
    cls = {"info":"t-blue","warn":"t-warn","ok":"","dim":"t-dim"}.get(level,"")
    return f'<span class="t-dim">[{ts}]</span> <span class="{cls}">{msg}</span>\n'

def priority_score(desc: str) -> str:
    desc_lower = desc.lower()
    high_kw = ["payment","checkout","login","auth","purchase","order","credit"]
    med_kw  = ["filter","search","navigation","display","load"]
    if any(k in desc_lower for k in high_kw): return "🔴 High"
    if any(k in desc_lower for k in med_kw):  return "🟡 Medium"
    return "🟢 Low"

# ─────────────────────────────────────────────────────────────────
#  MOCK AI ENGINE  (replaces CrewAI/OpenAI so app runs standalone)
#  Swap generate_test_cases() with your real crew when deploying.
# ─────────────────────────────────────────────────────────────────
def generate_test_cases(
    user_story: str,
    prefix: str = "CT",
    scope: str = "Full",
    num_cases: int = 8,
    progress_cb=None,
) -> pd.DataFrame:
    """
    Production hook – replace body with your CrewAI crew.kickoff() call.
    Currently generates realistic-looking test cases from keywords.
    """
    # ── keyword extraction ────────────────────
    words = re.findall(r'\b\w{4,}\b', user_story.lower())
    feature_kw = [w for w in words if w not in {
        "user","that","with","this","should","when","from","have",
        "able","will","must","need","page","view","the","and","for"
    }][:6]
    feature = " ".join(feature_kw[:3]).title() or "Feature"

    templates = [
        ("Verify {f} loads successfully",
         "1. Navigate to the {f} page.\n2. Observe page rendering.",
         "Page loads within 3s with all elements visible.", "User is logged in", "Page is displayed correctly"),
        ("Verify {f} with valid input",
         "1. Enter valid data in all fields.\n2. Submit the form.",
         "Success message shown; data is saved correctly.", "User on {f} page", "Record created"),
        ("Verify {f} with invalid input",
         "1. Enter invalid/empty data.\n2. Submit the form.",
         "Error messages displayed; data NOT saved.", "User on {f} page", "Form remains open"),
        ("Verify {f} for mobile viewport",
         "1. Open {f} on mobile (375px).\n2. Interact with all controls.",
         "Layout responsive; no overflow.", "Mobile device / emulator", "UI matches design spec"),
        ("Verify {f} error state handling",
         "1. Simulate network error during {f}.\n2. Observe behaviour.",
         "Friendly error message; no crash.", "API mocked to fail", "Error state resolved gracefully"),
        ("Verify {f} accessibility (WCAG 2.1)",
         "1. Run axe scan on {f}.\n2. Test keyboard navigation.",
         "Zero critical a11y violations.", "Accessibility tool available", "Compliance confirmed"),
        ("Verify {f} performance (LCP < 2.5s)",
         "1. Open Lighthouse on {f}.\n2. Record LCP metric.",
         "LCP < 2.5 seconds.", "Production-like env", "Performance budget met"),
        ("Verify {f} with concurrent users",
         "1. Simulate 50 concurrent sessions on {f}.\n2. Monitor response times.",
         "P95 latency < 1s, no 5xx errors.", "Load-testing env ready", "System stable"),
        ("Verify {f} after session timeout",
         "1. Remain idle 30 min.\n2. Attempt action on {f}.",
         "Redirected to login; no data loss.", "User logged in", "Session refreshed on re-login"),
        ("Verify {f} data persistence after refresh",
         "1. Input data in {f}.\n2. Hard-refresh browser.",
         "Data retained from server/cache.", "Valid session", "Data matches pre-refresh state"),
    ]

    scope_map = {"Quick (Smoke)": 3, "Standard": 6, "Full": num_cases, "Regression": num_cases + 4}
    count = scope_map.get(scope, num_cases)
    rows = []
    for i in range(min(count, len(templates))):
        if progress_cb:
            progress_cb(int((i / count) * 85))
        tmpl = templates[i]
        desc   = tmpl[0].replace("{f}", feature)
        steps  = tmpl[1].replace("{f}", feature)
        expect = tmpl[2]
        pre    = tmpl[3].replace("{f}", feature)
        post   = tmpl[4]
        rows.append({
            "CT Test Case IDs":     make_ct_id(prefix, i + 1),
            "OSD Number":           f"OSD-{1000 + i}",
            "ACC FLOW ID":          f"AF-{200 + i}",
            "ACC":                  f"ACC-{10 + i}",
            "Test Case Description": desc,
            "Test case steps":      steps,
            "Expected results":     expect,
            "Pre-condition":        pre,
            "Post condition":       post,
            "Dev Status":           "Not Started",
            "Testing Status":       "Not Started",
            "Comments":             "",
            "ETA of Dev":           "TBD",
            "ETA of CT":            "TBD",
            "TESTER":               "",
        })
    if progress_cb:
        progress_cb(100)
    return pd.DataFrame(rows, columns=COLUMNS)

# ─────────────────────────────────────────────
# ██  LOGIN PAGE
# ─────────────────────────────────────────────
def show_login():
    st.markdown(GLOBAL_CSS, unsafe_allow_html=True)
    st.markdown('<div class="qaforge-bg"></div>', unsafe_allow_html=True)
    st.markdown('<div class="login-shell">', unsafe_allow_html=True)

    # card wrapper — we fake the card via markdown + st widgets inside
    st.markdown("""
    <div class="login-card">
      <div class="login-logo">
        <span class="logo-icon">⚡</span>
        <h1>QA FORGE</h1>
        <p>AI-POWERED TEST CASE ENGINE</p>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # centre the form
    _, col, _ = st.columns([1, 1.6, 1])
    with col:
        st.markdown("""
        <div style="background:linear-gradient(135deg,rgba(12,18,34,.97),rgba(17,24,39,.97));
                    border:1px solid #1e2d47; border-radius:20px; padding:2.5rem 2.2rem;
                    box-shadow:0 40px 80px rgba(0,0,0,.6),inset 0 1px 0 rgba(255,255,255,.04);
                    position:relative; overflow:hidden;">
          <div style="position:absolute;top:0;left:0;right:0;height:2px;
                      background:linear-gradient(90deg,transparent,#00d4ff,#7c3aed,transparent);
                      animation:shimmer 3s linear infinite;"></div>
          <div style="text-align:center;margin-bottom:2rem;">
            <div style="font-size:3rem;filter:drop-shadow(0 0 16px rgba(0,212,255,.8));">⚡</div>
            <div style="font-family:Orbitron,monospace;font-size:1.7rem;font-weight:900;
                        background:linear-gradient(90deg,#00d4ff,#7c3aed);
                        -webkit-background-clip:text;-webkit-text-fill-color:transparent;
                        letter-spacing:2px;margin:.3rem 0 .1rem;">QA FORGE</div>
            <div style="font-size:.72rem;color:#64748b;font-family:'Space Mono',monospace;letter-spacing:1px;">
              AI-POWERED TEST CASE ENGINE</div>
          </div>
        """, unsafe_allow_html=True)

        username = st.text_input("USERNAME", key="li_user", placeholder="Enter username")
        password = st.text_input("PASSWORD", key="li_pass", placeholder="••••••••", type="password")

        if st.button("⚡  ACCESS SYSTEM", key="li_btn"):
            if st.session_state.login_attempts >= 5:
                st.markdown('<div class="login-error">🔒 Account locked – too many attempts</div>',
                            unsafe_allow_html=True)
            elif username in USERS and _hash(password) == USERS[username]["password"]:
                st.session_state.authenticated = True
                st.session_state.username = username
                st.session_state.login_attempts = 0
                st.rerun()
            else:
                st.session_state.login_attempts += 1
                remain = 5 - st.session_state.login_attempts
                st.markdown(
                    f'<div class="login-error">❌ Invalid credentials — {remain} attempt(s) remaining</div>',
                    unsafe_allow_html=True)

        st.markdown("""
          <div class="demo-hint">
            Demo credentials<br>
            user: <span>admin</span> &nbsp;|&nbsp; pass: <span>admin123</span><br>
            user: <span>demo</span> &nbsp;&nbsp;|&nbsp; pass: <span>demo</span>
          </div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────
# ██  MAIN APP
# ─────────────────────────────────────────────
def show_app():
    st.markdown(GLOBAL_CSS, unsafe_allow_html=True)
    st.markdown('<div class="qaforge-bg"></div>', unsafe_allow_html=True)

    # ── top navbar ────────────────────────────
    user_info = USERS[st.session_state.username]
    st.markdown(f"""
    <div class="top-nav">
      <div class="nav-logo">⚡ QA FORGE</div>
      <div class="nav-right">
        <span class="nav-badge">🟢 LIVE</span>
        <span class="nav-badge fpill-purple" style="border-color:rgba(124,58,237,.3);color:#a78bfa;background:rgba(124,58,237,.07);">
          {user_info['role']}
        </span>
        <span class="nav-user">👤 {user_info['name']}</span>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── stat row ──────────────────────────────
    hist_count  = len(st.session_state.history)
    total_tc    = sum(len(h["df"]) for h in st.session_state.history) if st.session_state.history else 0
    last_run    = st.session_state.history[-1]["ts"] if st.session_state.history else "—"

    st.markdown(f"""
    <div class="app-shell">
    <div class="stat-row">
      <div class="stat-card">
        <div class="stat-num">{hist_count}</div>
        <div class="stat-label">Sessions Run</div>
      </div>
      <div class="stat-card">
        <div class="stat-num">{total_tc}</div>
        <div class="stat-label">Test Cases Generated</div>
      </div>
      <div class="stat-card">
        <div class="stat-num">{len(st.session_state.history[-1]["df"]) if st.session_state.history else 0}</div>
        <div class="stat-label">Last Batch Size</div>
      </div>
      <div class="stat-card">
        <div class="stat-num" style="font-size:1.1rem;">{last_run if last_run == "—" else last_run[:16]}</div>
        <div class="stat-label">Last Run</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── main tabs ─────────────────────────────
    tab_gen, tab_edit, tab_hist, tab_analytics = st.tabs(
        ["⚡  Generator", "✏️  Editor & Export", "📂  History", "📊  Analytics"]
    )

    # ══════════════════════════════════════════
    # TAB 1 – GENERATOR
    # ══════════════════════════════════════════
    with tab_gen:
        st.markdown('<div class="section-header"><div class="dot"></div>CONFIGURATION</div>',
                    unsafe_allow_html=True)

        cfg_col1, cfg_col2, cfg_col3 = st.columns([1.4, 1, 1])
        with cfg_col1:
            ct_prefix = st.text_input("CT ID Prefix", value="CT", help="Prefix for CT Test Case IDs (e.g. CT → CT-0001)")
        with cfg_col2:
            scope = st.selectbox("Test Scope", ["Quick (Smoke)", "Standard", "Full", "Regression"])
        with cfg_col3:
            num_cases = st.slider("No. of Test Cases", min_value=3, max_value=30, value=8)

        col_upload, col_opts = st.columns([1.5, 1])

        with col_upload:
            st.markdown('<div class="section-header"><div class="dot"></div>USER STORY INPUT</div>',
                        unsafe_allow_html=True)
            input_mode = st.radio("Input mode", ["📤 Upload File", "✍️ Paste Text"], horizontal=True)

            user_story_text = ""
            if input_mode == "📤 Upload File":
                uploaded = st.file_uploader(
                    "Upload User Story (Word / .txt)",
                    type=["docx", "txt"],
                    label_visibility="collapsed",
                )
                if uploaded:
                    user_story_text = read_uploaded(uploaded)
                    st.markdown(
                        f'<div class="banner banner-success">✅ Loaded: <b>{uploaded.name}</b>'
                        f' — {len(user_story_text):,} chars</div>',
                        unsafe_allow_html=True)
            else:
                user_story_text = st.text_area(
                    "Paste your user story here",
                    height=180,
                    placeholder="As a user, I want to...",
                    label_visibility="collapsed",
                )

            # optional extra context
            with st.expander("➕ Additional Context / Constraints"):
                extra_ctx = st.text_area("Extra context (test env, edge cases, exclusions):", height=100)

        with col_opts:
            st.markdown('<div class="section-header"><div class="dot"></div>AI OPTIONS</div>',
                        unsafe_allow_html=True)
            st.markdown("""
            <div class="feature-pills">
              <span class="fpill fpill-blue">🔵 Happy Path</span>
              <span class="fpill fpill-purple">🟣 Negative Cases</span>
              <span class="fpill fpill-amber">🟡 Edge Cases</span>
              <span class="fpill fpill-green">🟢 A11y</span>
            </div>
            """, unsafe_allow_html=True)
            include_negative   = st.checkbox("Include Negative Test Cases", value=True)
            include_edge       = st.checkbox("Include Edge Cases",          value=True)
            include_a11y       = st.checkbox("Include Accessibility Tests", value=False)
            include_perf       = st.checkbox("Include Performance Tests",   value=False)
            auto_priority      = st.checkbox("Auto-assign Priority",        value=True)
            deduplicate        = st.checkbox("Deduplicate Similar Cases",   value=True)

            st.markdown('<div class="section-header" style="margin-top:1rem;"><div class="dot"></div>EXPORT FORMAT</div>',
                        unsafe_allow_html=True)
            export_fmt = st.selectbox("Default export", ["Excel (.xlsx)", "CSV (.csv)", "JSON (.json)"])

        # ── Generate button ───────────────────
        st.markdown("<br>", unsafe_allow_html=True)
        gen_col, _ = st.columns([1, 3])
        with gen_col:
            run = st.button("⚡  GENERATE TEST CASES", key="gen_btn")

        if run:
            if not user_story_text.strip():
                st.markdown('<div class="banner banner-warn">⚠️ Please provide a user story first.</div>',
                            unsafe_allow_html=True)
            else:
                st.markdown('<div class="section-header"><div class="dot"></div>AI ENGINE LOG</div>',
                            unsafe_allow_html=True)

                log_placeholder = st.empty()
                prog_placeholder = st.empty()
                log_html = ""

                def append_log(msg, level="info"):
                    nonlocal log_html
                    log_html += log_line(msg, level)
                    log_placeholder.markdown(
                        f'<div class="terminal">{log_html}</div>', unsafe_allow_html=True)

                def set_prog(pct):
                    prog_placeholder.markdown(
                        f'<div class="prog-wrap"><div class="prog-fill" style="width:{pct}%"></div></div>',
                        unsafe_allow_html=True)

                append_log("Initialising QA Forge AI Engine …", "dim")
                time.sleep(0.3)
                append_log(f"Scope: {scope}  |  Prefix: {ct_prefix}  |  Target cases: {num_cases}")
                time.sleep(0.2)
                append_log("Parsing user story …", "dim")
                time.sleep(0.3)
                append_log(f"Story length: {len(user_story_text)} chars — keywords extracted ✓", "ok")
                time.sleep(0.2)
                append_log("Dispatching to Senior QA Agent …")
                time.sleep(0.3)

                df = generate_test_cases(
                    user_story_text + ("\n\n" + extra_ctx if extra_ctx else ""),
                    prefix=ct_prefix,
                    scope=scope,
                    num_cases=num_cases,
                    progress_cb=set_prog,
                )

                append_log(f"Agent returned {len(df)} raw test cases ✓", "ok")
                time.sleep(0.2)

                if deduplicate and not df.empty:
                    before = len(df)
                    df = df.drop_duplicates(subset=["Test Case Description"])
                    removed = before - len(df)
                    append_log(f"Deduplication: removed {removed} duplicate(s)", "warn" if removed else "dim")

                if auto_priority and not df.empty:
                    df.insert(4, "Priority",
                               df["Test Case Description"].apply(priority_score))
                    append_log("Priority scoring applied ✓", "ok")

                append_log("Formatting complete — ready for download.", "ok")
                set_prog(100)
                time.sleep(0.3)

                # ── store in history ──────────────
                session_id = str(uuid.uuid4())[:8].upper()
                ts_now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
                snippet = (user_story_text[:60] + "…") if len(user_story_text) > 60 else user_story_text
                st.session_state.history.append({
                    "id": session_id, "name": snippet, "ts": ts_now,
                    "df": df.copy(), "config": {
                        "prefix": ct_prefix, "scope": scope,
                        "num": num_cases, "fmt": export_fmt
                    }
                })
                st.session_state.active_df = df.copy()

                st.markdown(
                    f'<div class="banner banner-success">✅ Generated <b>{len(df)}</b> test cases'
                    f' — Session ID: <b>{session_id}</b></div>',
                    unsafe_allow_html=True)

                # ── preview table ─────────────────
                st.markdown('<div class="section-header"><div class="dot"></div>PREVIEW</div>',
                            unsafe_allow_html=True)
                preview_cols = ["CT Test Case IDs","Test Case Description","Expected results","Dev Status","Testing Status"]
                if "Priority" in df.columns:
                    preview_cols.insert(1, "Priority")
                st.dataframe(df[preview_cols], use_container_width=True, height=320)
                st.info("💡 Head to the **Editor & Export** tab to edit, filter, and download the full dataset.")

    # ══════════════════════════════════════════
    # TAB 2 – EDITOR & EXPORT
    # ══════════════════════════════════════════
    with tab_edit:
        df_work = st.session_state.active_df

        if df_work is None or df_work.empty:
            st.markdown(
                '<div class="banner banner-info">ℹ️ No test cases loaded. Run the Generator tab first.</div>',
                unsafe_allow_html=True)
        else:
            st.markdown('<div class="section-header"><div class="dot"></div>FILTERS</div>',
                        unsafe_allow_html=True)

            f1, f2, f3 = st.columns(3)
            with f1:
                flt_status = st.multiselect("Testing Status", df_work["Testing Status"].unique().tolist(),
                                             default=df_work["Testing Status"].unique().tolist())
            with f2:
                flt_dev = st.multiselect("Dev Status", df_work["Dev Status"].unique().tolist(),
                                          default=df_work["Dev Status"].unique().tolist())
            with f3:
                if "Priority" in df_work.columns:
                    flt_pri = st.multiselect("Priority", df_work["Priority"].unique().tolist(),
                                              default=df_work["Priority"].unique().tolist())
                else:
                    flt_pri = None

            mask = (
                df_work["Testing Status"].isin(flt_status) &
                df_work["Dev Status"].isin(flt_dev)
            )
            if flt_pri is not None and "Priority" in df_work.columns:
                mask &= df_work["Priority"].isin(flt_pri)
            df_filtered = df_work[mask].copy()

            st.markdown(
                f'<div class="banner banner-info">Showing <b>{len(df_filtered)}</b> of <b>{len(df_work)}</b> test cases</div>',
                unsafe_allow_html=True)

            # ── inline editable grid ──────────────
            st.markdown('<div class="section-header"><div class="dot"></div>INLINE EDITOR</div>',
                        unsafe_allow_html=True)
            edited_df = st.data_editor(
                df_filtered, use_container_width=True, height=400,
                column_config={
                    "Dev Status": st.column_config.SelectboxColumn(
                        options=["Not Started","In Progress","Done","Blocked"]),
                    "Testing Status": st.column_config.SelectboxColumn(
                        options=["Not Started","In Progress","Pass","Fail","Blocked","N/A"]),
                    "ETA of Dev": st.column_config.TextColumn(default="TBD"),
                    "ETA of CT":  st.column_config.TextColumn(default="TBD"),
                    "TESTER":     st.column_config.TextColumn(default=""),
                },
                num_rows="dynamic",
                key="editor_grid",
            )

            # persist back
            for idx in edited_df.index:
                for col in edited_df.columns:
                    st.session_state.active_df.at[idx, col] = edited_df.at[idx, col]

            # ── export buttons ────────────────────
            st.markdown('<div class="section-header" style="margin-top:1.5rem;"><div class="dot"></div>EXPORT</div>',
                        unsafe_allow_html=True)
            e1, e2, e3, e4 = st.columns(4)
            final_df = st.session_state.active_df

            with e1:
                st.download_button("📥 Download Excel",
                    data=df_to_excel_bytes(final_df),
                    file_name=f"QAForge_TestCases_{datetime.date.today()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with e2:
                st.download_button("📄 Download CSV",
                    data=final_df.to_csv(index=False).encode(),
                    file_name=f"QAForge_TestCases_{datetime.date.today()}.csv",
                    mime="text/csv")
            with e3:
                st.download_button("🔧 Download JSON",
                    data=final_df.to_json(orient="records", indent=2).encode(),
                    file_name=f"QAForge_TestCases_{datetime.date.today()}.json",
                    mime="application/json")
            with e4:
                st.download_button("📋 Markdown Report",
                    data=final_df.to_markdown(index=False).encode(),
                    file_name=f"QAForge_Report_{datetime.date.today()}.md",
                    mime="text/markdown")

    # ══════════════════════════════════════════
    # TAB 3 – HISTORY
    # ══════════════════════════════════════════
    with tab_hist:
        if not st.session_state.history:
            st.markdown(
                '<div class="banner banner-info">ℹ️ No history yet — generate some test cases first.</div>',
                unsafe_allow_html=True)
        else:
            st.markdown('<div class="section-header"><div class="dot"></div>PREVIOUS SESSIONS</div>',
                        unsafe_allow_html=True)
            for h in reversed(st.session_state.history):
                h_col1, h_col2 = st.columns([3, 1])
                with h_col1:
                    st.markdown(f"""
                    <div class="hist-card">
                      <div class="hist-meta">📅 {h['ts']} &nbsp;|&nbsp; 🆔 {h['id']} &nbsp;|&nbsp; 📋 {len(h['df'])} cases</div>
                      <div class="hist-title">{h['name']}</div>
                      <div class="hist-meta" style="margin-top:.3rem;">
                        Prefix: {h['config']['prefix']} &nbsp;|&nbsp; Scope: {h['config']['scope']}
                      </div>
                    </div>
                    """, unsafe_allow_html=True)
                with h_col2:
                    if st.button(f"Load →", key=f"load_{h['id']}"):
                        st.session_state.active_df = h["df"].copy()
                        st.success(f"Loaded session {h['id']}")
                    st.download_button("⬇ Excel", data=df_to_excel_bytes(h["df"]),
                                       file_name=f"QAForge_{h['id']}.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                       key=f"dl_{h['id']}")

            if st.button("🗑️  Clear All History", type="secondary"):
                st.session_state.history = []
                st.session_state.active_df = None
                st.rerun()

    # ══════════════════════════════════════════
    # TAB 4 – ANALYTICS
    # ══════════════════════════════════════════
    with tab_analytics:
        all_dfs = [h["df"] for h in st.session_state.history] if st.session_state.history else []
        if not all_dfs:
            st.markdown(
                '<div class="banner banner-info">ℹ️ No data to analyse yet.</div>',
                unsafe_allow_html=True)
        else:
            big_df = pd.concat(all_dfs, ignore_index=True)

            # ── KPI bar ───────────────────────────
            kpi1, kpi2, kpi3, kpi4 = st.columns(4)
            pass_pct = round((big_df["Testing Status"] == "Pass").sum() / max(len(big_df), 1) * 100, 1)
            fail_pct = round((big_df["Testing Status"] == "Fail").sum() / max(len(big_df), 1) * 100, 1)
            done_pct = round((big_df["Dev Status"] == "Done").sum() / max(len(big_df), 1) * 100, 1)
            kpi1.metric("Total Test Cases", len(big_df))
            kpi2.metric("Pass Rate",  f"{pass_pct}%")
            kpi3.metric("Fail Rate",  f"{fail_pct}%")
            kpi4.metric("Dev Done",   f"{done_pct}%")

            st.markdown("<br>", unsafe_allow_html=True)
            a1, a2 = st.columns(2)

            # ── Testing Status breakdown ──────────
            with a1:
                ts_counts = big_df["Testing Status"].value_counts()
                bar_rows = ""
                colors = {"Pass":"#10b981","Fail":"#ef4444","Not Started":"#64748b",
                          "In Progress":"#00d4ff","Blocked":"#f59e0b","N/A":"#a78bfa"}
                total = len(big_df)
                for status, cnt in ts_counts.items():
                    pct = int(cnt / total * 100)
                    c   = colors.get(status, "#64748b")
                    bar_rows += f"""
                    <div class="bar-row">
                      <div class="bar-label">{status}</div>
                      <div class="bar-track">
                        <div class="bar-fill" style="width:{pct}%;background:{c};"></div>
                      </div>
                      <span style="font-size:.72rem;color:#64748b;font-family:'Space Mono',monospace;width:30px;">{cnt}</span>
                    </div>"""
                st.markdown(f"""
                <div class="chart-card">
                  <h4>Testing Status Distribution</h4>
                  {bar_rows}
                </div>""", unsafe_allow_html=True)

            # ── Dev Status breakdown ──────────────
            with a2:
                ds_counts = big_df["Dev Status"].value_counts()
                bar_rows2 = ""
                for status, cnt in ds_counts.items():
                    pct = int(cnt / total * 100)
                    c   = colors.get(status, "#64748b")
                    bar_rows2 += f"""
                    <div class="bar-row">
                      <div class="bar-label">{status}</div>
                      <div class="bar-track">
                        <div class="bar-fill" style="width:{pct}%;background:{c};"></div>
                      </div>
                      <span style="font-size:.72rem;color:#64748b;font-family:'Space Mono',monospace;width:30px;">{cnt}</span>
                    </div>"""
                st.markdown(f"""
                <div class="chart-card">
                  <h4>Dev Status Distribution</h4>
                  {bar_rows2}
                </div>""", unsafe_allow_html=True)

            # ── session history line ──────────────
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown('<div class="section-header"><div class="dot"></div>GENERATION HISTORY</div>',
                        unsafe_allow_html=True)
            hist_chart_data = pd.DataFrame([
                {"Session": h["id"], "Test Cases": len(h["df"]), "Timestamp": h["ts"]}
                for h in st.session_state.history
            ])
            st.bar_chart(hist_chart_data.set_index("Session")["Test Cases"])

            # ── Priority heatmap (if exists) ──────
            if "Priority" in big_df.columns:
                st.markdown('<div class="section-header"><div class="dot"></div>PRIORITY BREAKDOWN</div>',
                            unsafe_allow_html=True)
                pri_counts = big_df["Priority"].value_counts()
                p1, p2, p3 = st.columns(3)
                for col, emoji_label in zip([p1, p2, p3],
                                            ["🔴 High", "🟡 Medium", "🟢 Low"]):
                    cnt = pri_counts.get(emoji_label, 0)
                    col.metric(emoji_label, cnt)

    # close app-shell div
    st.markdown("</div>", unsafe_allow_html=True)

    # ── sidebar: logout + quick stats ─────────
    with st.sidebar:
        st.markdown("""
        <div style="font-family:'Orbitron',monospace;font-size:.95rem;font-weight:700;
                    color:#00d4ff;letter-spacing:2px;padding:.5rem 0 1rem;">⚡ QA FORGE</div>
        """, unsafe_allow_html=True)
        st.markdown(f"**Role:** {user_info['role']}")
        st.markdown(f"**User:** {user_info['name']}")
        st.divider()
        st.markdown("**Quick Stats**")
        st.metric("Sessions", len(st.session_state.history))
        st.metric("Total Cases", sum(len(h["df"]) for h in st.session_state.history))
        st.divider()
        if st.button("🔓 Logout", use_container_width=True):
            for k in ["authenticated", "username", "active_df"]:
                st.session_state[k] = False if k == "authenticated" else None if k == "active_df" else ""
            st.rerun()

# ─────────────────────────────────────────────
# ROUTER
# ─────────────────────────────────────────────
if not st.session_state.authenticated:
    show_login()
else:
    show_app()
