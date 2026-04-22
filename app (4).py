# ============================================================
#  QA·AI  —  Ultimate Hybrid Test Case Generator
#  Streamlit · SQLite · Anthropic Claude API
#  15 Test Cases | AI + Smart Dynamic Fallback | Full Auth
# ============================================================

from __future__ import annotations

import streamlit as st
import sqlite3
import pandas as pd
import hashlib
import io
import re
import json
import os
from datetime import datetime
from pathlib import Path

# ── Optional imports ──
try:
    import PyPDF2
except ImportError:
    PyPDF2 = None

try:
    from docx import Document as DocxDocument
except ImportError:
    DocxDocument = None

try:
    import anthropic
except ImportError:
    anthropic = None

try:
    import openpyxl
except ImportError:
    openpyxl = None

# ─────────────────────────────────────────────────────────────
#  PAGE CONFIG  (must be first Streamlit call)
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="QA·AI — Ultimate Hybrid Engine",
    page_icon="🧬",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────
#  GLOBAL CSS  — dark-glass design, electric-teal accents
# ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&family=DM+Sans:wght@300;400;500;600;700&display=swap');

:root {
    --bg:        #0b0f1a;
    --surface:   #111827;
    --surface2:  #1a2236;
    --border:    #1f2d45;
    --teal:      #00e5c3;
    --teal-dim:  rgba(0,229,195,0.12);
    --teal-glow: 0 0 24px rgba(0,229,195,0.25);
    --amber:     #f59e0b;
    --red:       #ef4444;
    --green:     #22c55e;
    --text:      #e2e8f0;
    --muted:     #64748b;
    --radius:    12px;
}

html, body, [data-testid="stAppViewContainer"] {
    background: var(--bg) !important;
    color: var(--text) !important;
    font-family: 'DM Sans', sans-serif !important;
}

[data-testid="stSidebar"] {
    background: var(--surface) !important;
    border-right: 1px solid var(--border) !important;
}
[data-testid="stSidebar"] * { color: var(--text) !important; }

h1, h2, h3 {
    font-family: 'Space Mono', monospace !important;
    color: var(--text) !important;
}

[data-testid="stMetric"] {
    background: var(--surface2) !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--radius) !important;
    padding: 16px 20px !important;
}
[data-testid="stMetricLabel"] { color: var(--muted) !important; font-size: 0.78rem !important; }
[data-testid="stMetricValue"] { color: var(--teal) !important; font-family: 'Space Mono', monospace !important; }

[data-testid="stTextInput"] input,
[data-testid="stSelectbox"] div[data-baseweb="select"] > div,
[data-testid="stFileUploader"] {
    background: var(--surface2) !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--radius) !important;
    color: var(--text) !important;
}
[data-testid="stTextInput"] input:focus {
    border-color: var(--teal) !important;
    box-shadow: var(--teal-glow) !important;
}

[data-testid="stButton"] > button[kind="primary"],
[data-testid="stFormSubmitButton"] > button {
    background: var(--teal) !important;
    color: #0b0f1a !important;
    font-family: 'Space Mono', monospace !important;
    font-weight: 700 !important;
    border: none !important;
    border-radius: var(--radius) !important;
    padding: 10px 26px !important;
    letter-spacing: 0.05em !important;
    transition: filter .2s, transform .15s !important;
}
[data-testid="stButton"] > button[kind="primary"]:hover,
[data-testid="stFormSubmitButton"] > button:hover {
    filter: brightness(1.15) !important;
    transform: translateY(-2px) !important;
}

[data-testid="stButton"] > button[kind="secondary"] {
    background: var(--surface2) !important;
    color: var(--teal) !important;
    border: 1px solid var(--teal) !important;
    border-radius: var(--radius) !important;
    font-family: 'Space Mono', monospace !important;
    transition: background .2s !important;
}
[data-testid="stButton"] > button[kind="secondary"]:hover {
    background: var(--teal-dim) !important;
}

[data-testid="stDataFrame"] {
    border: 1px solid var(--border) !important;
    border-radius: var(--radius) !important;
    overflow: hidden !important;
}

[data-testid="stAlert"] {
    border-radius: var(--radius) !important;
    border: 1px solid var(--border) !important;
}

[data-testid="stTabs"] [data-baseweb="tab-list"] {
    background: var(--surface2) !important;
    border-radius: var(--radius) !important;
    padding: 4px !important;
}
[data-testid="stTabs"] [data-baseweb="tab"] {
    color: var(--muted) !important;
    font-family: 'Space Mono', monospace !important;
    font-size: 0.82rem !important;
}
[data-testid="stTabs"] [aria-selected="true"] {
    background: var(--teal-dim) !important;
    color: var(--teal) !important;
    border-radius: 8px !important;
}

hr { border-color: var(--border) !important; }
[data-testid="stSpinner"] { color: var(--teal) !important; }

[data-testid="stExpander"] {
    background: var(--surface2) !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--radius) !important;
}

.mode-badge {
    display: inline-block;
    padding: 4px 12px;
    border-radius: 20px;
    font-family: 'Space Mono', monospace;
    font-size: 0.75rem;
    font-weight: 700;
    letter-spacing: 0.08em;
}
.mode-ai { background: rgba(0,229,195,0.15); color: #00e5c3; border: 1px solid #00e5c3; }
.mode-fallback { background: rgba(245,158,11,0.15); color: #f59e0b; border: 1px solid #f59e0b; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
#  DATABASE
# ─────────────────────────────────────────────────────────────
DB_PATH = Path("qa_ai.db")

@st.cache_resource
def get_db() -> sqlite3.Connection:
    conn = sqlite3.connect(str(DB_PATH), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    _create_schema(conn)
    return conn

def _create_schema(conn: sqlite3.Connection):
    conn.executescript("""
    CREATE TABLE IF NOT EXISTS users (
        username TEXT PRIMARY KEY,
        password TEXT NOT NULL,
        created_at TEXT DEFAULT (datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS teams (
        id   INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        owner TEXT NOT NULL,
        created_at TEXT DEFAULT (datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS team_members (
        username TEXT NOT NULL,
        team_id  INTEGER NOT NULL,
        PRIMARY KEY (username, team_id)
    );
    CREATE TABLE IF NOT EXISTS projects (
        id       INTEGER PRIMARY KEY AUTOINCREMENT,
        name     TEXT NOT NULL,
        team_id  INTEGER NOT NULL,
        created_at TEXT DEFAULT (datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS testcases (
        id         INTEGER PRIMARY KEY AUTOINCREMENT,
        username   TEXT NOT NULL,
        project_id INTEGER NOT NULL,
        story      TEXT NOT NULL,
        result     TEXT NOT NULL,
        engine     TEXT DEFAULT 'fallback',
        created_at TEXT DEFAULT (datetime('now'))
    );
    """)
    # Add engine column if upgrading from older DB
    try:
        conn.execute("ALTER TABLE testcases ADD COLUMN engine TEXT DEFAULT 'fallback'")
        conn.commit()
    except Exception:
        pass
    conn.commit()

# ─────────────────────────────────────────────────────────────
#  AUTH
# ─────────────────────────────────────────────────────────────
def _hash(password: str) -> str:
    return hashlib.sha256(password.encode()).hexdigest()

def register_user(username: str, password: str) -> tuple[bool, str]:
    if len(username) < 3:
        return False, "Username must be at least 3 characters."
    if len(password) < 6:
        return False, "Password must be at least 6 characters."
    db = get_db()
    try:
        db.execute("INSERT INTO users (username, password) VALUES (?,?)",
                   (username.strip(), _hash(password)))
        db.commit()
        return True, "Account created successfully."
    except sqlite3.IntegrityError:
        return False, "Username already taken."

def login_user(username: str, password: str) -> bool:
    db = get_db()
    row = db.execute(
        "SELECT 1 FROM users WHERE username=? AND password=?",
        (username.strip(), _hash(password))
    ).fetchone()
    return row is not None

# ─────────────────────────────────────────────────────────────
#  FILE READER
# ─────────────────────────────────────────────────────────────
def read_uploaded_file(file) -> str:
    name = file.name.lower()
    try:
        if name.endswith(".txt"):
            return file.read().decode("utf-8", errors="replace")
        elif name.endswith(".pdf"):
            if PyPDF2 is None:
                return "[PyPDF2 not installed — cannot read PDF]"
            reader = PyPDF2.PdfReader(file)
            return "\n".join(p.extract_text() or "" for p in reader.pages)
        elif name.endswith(".docx"):
            if DocxDocument is None:
                return "[python-docx not installed — cannot read DOCX]"
            doc = DocxDocument(file)
            return "\n".join(p.text for p in doc.paragraphs)
    except Exception as exc:
        return f"[Error reading file: {exc}]"
    return ""

# ─────────────────────────────────────────────────────────────
#  ULTRA-DYNAMIC FALLBACK ENGINE (15 cases, context-aware)
# ─────────────────────────────────────────────────────────────
def get_ultra_dynamic_fallback(story: str) -> list[dict]:
    """
    Parses the story to extract action verbs and target nouns,
    then generates 15 contextually relevant test cases — no dummy data.
    """
    words = re.findall(r'\b\w{4,}\b', story)

    # Detect action keyword from story
    action_keywords = ["login", "logout", "register", "upload", "download", "delete",
                       "update", "create", "search", "reset", "verify", "submit",
                       "edit", "view", "export", "import", "filter", "sort", "pay"]
    action = "Action"
    for w in words:
        if w.lower() in action_keywords:
            action = w.capitalize()
            break
    if action == "Action" and words:
        action = words[0].capitalize()

    # Detect target (first meaningful noun after the action)
    target = words[1].capitalize() if len(words) > 1 else "System"

    templates = [
        # Functional (5)
        {"id": "TC_01", "type": "Functional",    "priority": "High",
         "scen": f"Verify successful {action} of {target} with valid credentials",
         "pre":  f"User has a valid account and {target} feature is enabled",
         "steps": f"1. Navigate to {target}\n2. Enter valid data\n3. Trigger {action}\n4. Confirm response",
         "exp":  f"{target} {action} completes successfully with a confirmation message."},

        {"id": "TC_02", "type": "Functional",    "priority": "High",
         "scen": f"Validate {target} data is saved correctly after {action}",
         "pre":  f"{target} module is accessible and DB is reachable",
         "steps": f"1. Perform {action} on {target}\n2. Query database\n3. Compare stored vs input data",
         "exp":  f"Data persisted in {target} matches the input exactly without corruption."},

        {"id": "TC_03", "type": "Functional",    "priority": "High",
         "scen": f"Verify {action} redirects user to correct {target} landing page",
         "pre":  "User is authenticated",
         "steps": f"1. Initiate {action}\n2. Observe URL and page content after redirect",
         "exp":  f"User is redirected to the expected {target} page with correct content."},

        {"id": "TC_04", "type": "Functional",    "priority": "Medium",
         "scen": f"Check {target} updates in real-time after {action}",
         "pre":  f"{target} UI is loaded and network is stable",
         "steps": f"1. Perform {action}\n2. Observe {target} state without page refresh",
         "exp":  f"{target} reflects updated state immediately without manual refresh."},

        {"id": "TC_05", "type": "Functional",    "priority": "Medium",
         "scen": f"Confirm {action} triggers correct notification or email for {target}",
         "pre":  "Notification service is configured",
         "steps": f"1. Perform {action} on {target}\n2. Check email/notification inbox",
         "exp":  f"A notification is delivered within 60 seconds with correct {target} details."},

        # Negative (3)
        {"id": "TC_06", "type": "Negative",      "priority": "High",
         "scen": f"Test {action} {target} with empty/null required fields",
         "pre":  f"{target} form is accessible",
         "steps": f"1. Leave required fields blank\n2. Attempt {action}\n3. Observe system response",
         "exp":  "Inline validation errors appear; action is blocked without data loss."},

        {"id": "TC_07", "type": "Negative",      "priority": "High",
         "scen": f"Attempt {action} on {target} with expired or invalid session",
         "pre":  "Session token is manually expired",
         "steps": f"1. Expire token\n2. Attempt {action} on {target}\n3. Observe response",
         "exp":  f"System rejects the request and prompts user to re-authenticate."},

        {"id": "TC_08", "type": "Negative",      "priority": "Medium",
         "scen": f"Test {target} behavior when {action} is repeated in rapid succession",
         "pre":  "Rate limiter may or may not be configured",
         "steps": f"1. Trigger {action} on {target} 10x within 5 seconds\n2. Observe system",
         "exp":  "System handles gracefully — either rate-limits or processes idempotently."},

        # Security (2)
        {"id": "TC_09", "type": "Security",      "priority": "High",
         "scen": f"Verify {target} {action} is inaccessible to unauthorized roles",
         "pre":  "Two accounts exist: admin and guest",
         "steps": f"1. Log in as guest\n2. Attempt {action} on {target}\n3. Check HTTP response",
         "exp":  "403 Forbidden is returned; no data is exposed to unauthorized user."},

        {"id": "TC_10", "type": "Security",      "priority": "High",
         "scen": f"Check {target} input fields for SQL injection and XSS vulnerabilities",
         "pre":  f"{target} form with text inputs is rendered",
         "steps": f"1. Enter ' OR 1=1 -- in {target} fields\n2. Enter <script>alert(1)</script>\n3. Submit",
         "exp":  "Inputs are sanitised; no script executes and no SQL error leaks to UI."},

        # Boundary (2)
        {"id": "TC_11", "type": "Boundary",      "priority": "Medium",
         "scen": f"Validate {target} {action} at maximum allowed character/data limit",
         "pre":  f"Max limit for {target} fields is documented",
         "steps": f"1. Enter exactly max-limit data into {target}\n2. Perform {action}",
         "exp":  f"{action} succeeds at the exact boundary; exceeding it shows an error."},

        {"id": "TC_12", "type": "Boundary",      "priority": "Medium",
         "scen": f"Test {target} {action} with minimum boundary values (0 or 1)",
         "pre":  f"Minimum constraints for {target} are defined",
         "steps": f"1. Enter minimum valid value\n2. Perform {action}\n3. Observe response",
         "exp":  f"System accepts minimum value; below-minimum inputs are rejected cleanly."},

        # UX/UI (2)
        {"id": "TC_13", "type": "UX/UI",         "priority": "Medium",
         "scen": f"Verify {target} {action} UI is consistent across desktop and mobile",
         "pre":  "Chrome DevTools device emulation available",
         "steps": f"1. Open {target} on desktop\n2. Switch to mobile viewport\n3. Perform {action}",
         "exp":  f"No layout breaks; all {target} controls are accessible on mobile."},

        {"id": "TC_14", "type": "UX/UI",         "priority": "Low",
         "scen": f"Check error message quality when {target} {action} fails",
         "pre":  "Network or validation failure can be simulated",
         "steps": f"1. Force {action} failure on {target}\n2. Read displayed error message",
         "exp":  "Error message is human-readable, actionable, and does not leak stack trace."},

        # Performance (1)
        {"id": "TC_15", "type": "Performance",   "priority": "Low",
         "scen": f"Measure {target} {action} response time under normal and peak load",
         "pre":  "Performance baseline defined; monitoring tool available",
         "steps": f"1. Perform {action} on {target} 50 times\n2. Record avg response time",
         "exp":  f"{target} {action} completes in < 2 seconds under normal; < 5s under peak."},
    ]

    return [
        {
            "id":              t["id"],
            "scenario":        t["scen"],
            "preconditions":   t["pre"],
            "steps":           t["steps"],
            "expected_result": t["exp"],
            "priority":        t["priority"],
            "complexity":      "High" if t["priority"] == "High" else "Medium" if t["priority"] == "Medium" else "Low",
            "type":            t["type"],
        }
        for t in templates
    ]

# ─────────────────────────────────────────────────────────────
#  AI ENGINE  — Claude API (15 cases)
# ─────────────────────────────────────────────────────────────
_SYSTEM_PROMPT = """You are a senior QA engineer. Given a user story, generate exactly 15 structured test cases.

Distribution: 5 Functional, 3 Negative, 2 Security, 2 Boundary, 2 UX/UI, 1 Performance.

Respond ONLY with a valid JSON array — no markdown, no prose, no code fences.
Each object must have these exact keys:
  "id"              : "TC_01" .. "TC_15"
  "scenario"        : one-sentence test scenario
  "preconditions"   : what must be true before the test
  "steps"           : numbered steps as a single string (use \\n between steps)
  "expected_result" : what the system should do
  "priority"        : one of High | Medium | Low
  "complexity"      : one of High | Medium | Low
  "type"            : one of Functional | Negative | Boundary | Performance | Security | UX/UI

Return ONLY the JSON array. 15 objects total."""

def generate_test_cases(story: str, api_key: str) -> tuple[pd.DataFrame, str]:
    """
    Returns (DataFrame of 15 test cases, engine_used).
    engine_used is 'Claude AI' or 'Smart Fallback'.
    """
    cases = None
    engine_used = "Smart Fallback"

    if anthropic and api_key:
        try:
            client = anthropic.Anthropic(api_key=api_key)
            message = client.messages.create(
                model="claude-opus-4-6",
                max_tokens=4096,
                system=_SYSTEM_PROMPT,
                messages=[{"role": "user", "content": story}],
            )
            raw = message.content[0].text.strip()
            raw = re.sub(r"^```[a-z]*\n?|```$", "", raw, flags=re.MULTILINE).strip()
            parsed = json.loads(raw)
            if isinstance(parsed, list) and len(parsed) >= 10:
                cases = parsed
                engine_used = "Claude AI"
            else:
                st.warning("AI returned fewer than 10 cases — switching to Smart Fallback.")
        except json.JSONDecodeError as e:
            st.warning(f"AI returned invalid JSON — using Smart Fallback. ({e})")
        except Exception as e:
            st.warning(f"AI call failed — using Smart Fallback. ({e})")

    if not cases:
        cases = get_ultra_dynamic_fallback(story)

    df = pd.DataFrame(cases).rename(columns={
        "id":              "Test Case ID",
        "scenario":        "Scenario",
        "preconditions":   "Pre-conditions",
        "steps":           "Steps",
        "expected_result": "Expected Result",
        "priority":        "Priority",
        "complexity":      "Complexity",
        "type":            "Type",
    })
    return df, engine_used

# ─────────────────────────────────────────────────────────────
#  TEAM & PROJECT HELPERS
# ─────────────────────────────────────────────────────────────
def create_team(name: str, owner: str) -> tuple[bool, str]:
    if not name.strip():
        return False, "Team name cannot be empty."
    db = get_db()
    db.execute("INSERT INTO teams (name, owner) VALUES (?,?)", (name.strip(), owner))
    team_id = db.execute("SELECT last_insert_rowid()").fetchone()[0]
    db.execute("INSERT OR IGNORE INTO team_members (username, team_id) VALUES (?,?)",
               (owner, team_id))
    db.commit()
    return True, f"Team '{name}' created."

def get_user_teams(username: str) -> list[dict]:
    db = get_db()
    rows = db.execute("""
        SELECT t.id, t.name, t.owner, t.created_at
        FROM teams t JOIN team_members tm ON t.id = tm.team_id
        WHERE tm.username = ? ORDER BY t.id DESC
    """, (username,)).fetchall()
    return [dict(r) for r in rows]

def create_project(name: str, team_id: int) -> tuple[bool, str]:
    if not name.strip():
        return False, "Project name cannot be empty."
    db = get_db()
    db.execute("INSERT INTO projects (name, team_id) VALUES (?,?)", (name.strip(), team_id))
    db.commit()
    return True, f"Project '{name}' created."

def get_team_projects(team_id: int) -> list[dict]:
    db = get_db()
    rows = db.execute(
        "SELECT id, name, created_at FROM projects WHERE team_id=? ORDER BY id DESC",
        (team_id,)
    ).fetchall()
    return [dict(r) for r in rows]

def get_project_history(project_id: int, limit: int = 20) -> list[dict]:
    db = get_db()
    rows = db.execute("""
        SELECT id, username, story, result, engine, created_at
        FROM testcases WHERE project_id=?
        ORDER BY id DESC LIMIT ?
    """, (project_id, limit)).fetchall()
    return [dict(r) for r in rows]

def get_stats(project_id: int) -> dict:
    db = get_db()
    r = db.execute("""
        SELECT COUNT(*) as runs, COUNT(DISTINCT username) as contributors
        FROM testcases WHERE project_id=?
    """, (project_id,)).fetchone()
    return {"runs": r["runs"], "contributors": r["contributors"]}

# ─────────────────────────────────────────────────────────────
#  EXCEL / CSV EXPORT
# ─────────────────────────────────────────────────────────────
def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    if openpyxl is None:
        return df.to_csv(index=False).encode()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Test Cases")
        ws = writer.sheets["Test Cases"]
        for col in ws.columns:
            max_len = max((len(str(cell.value or "")) for cell in col), default=10)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 60)
    return buf.getvalue()

# ─────────────────────────────────────────────────────────────
#  SESSION STATE BOOTSTRAP
# ─────────────────────────────────────────────────────────────
for _key, _default in [
    ("authenticated",  False),
    ("username",       ""),
    ("api_key",        ""),
    ("active_team",    None),
    ("active_project", None),
]:
    if _key not in st.session_state:
        st.session_state[_key] = _default

# ─────────────────────────────────────────────────────────────
#  SIDEBAR — AUTH + NAV
# ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style='padding:20px 0 8px;'>
      <span style='font-family:"Space Mono",monospace;font-size:1.5rem;
                   color:#00e5c3;letter-spacing:0.08em;'>QA·AI</span>
      <br>
      <span style='font-size:0.72rem;color:#64748b;letter-spacing:0.15em;
                   text-transform:uppercase;'>Ultimate Hybrid Engine</span>
    </div>
    <hr style='margin:8px 0 16px;'>
    """, unsafe_allow_html=True)

    # ── AUTH ──
    if not st.session_state["authenticated"]:
        auth_tab, reg_tab = st.tabs(["🔑  Login", "✨  Register"])

        with auth_tab:
            li_user = st.text_input("Username", key="li_user")
            li_pwd  = st.text_input("Password", type="password", key="li_pwd")
            if st.button("Login", key="btn_login", use_container_width=True, type="primary"):
                if login_user(li_user, li_pwd):
                    st.session_state["authenticated"] = True
                    st.session_state["username"]       = li_user.strip()
                    st.rerun()
                else:
                    st.error("Invalid credentials.")

        with reg_tab:
            reg_user = st.text_input("Username", key="reg_user")
            reg_pwd  = st.text_input("Password", type="password", key="reg_pwd")
            reg_pwd2 = st.text_input("Confirm Password", type="password", key="reg_pwd2")
            if st.button("Create Account", key="btn_reg", use_container_width=True, type="primary"):
                if reg_pwd != reg_pwd2:
                    st.error("Passwords do not match.")
                else:
                    ok, msg = register_user(reg_user, reg_pwd)
                    st.success(msg) if ok else st.error(msg)
        st.stop()

    # ── LOGGED-IN USER ──
    user = st.session_state["username"]
    st.markdown(f"""
    <div style='display:flex;align-items:center;gap:10px;padding:4px 0 12px;'>
      <div style='width:34px;height:34px;border-radius:50%;background:#00e5c3;
                  display:flex;align-items:center;justify-content:center;
                  font-family:"Space Mono",monospace;font-weight:700;color:#0b0f1a;'>
        {user[0].upper()}
      </div>
      <span style='font-size:0.9rem;'>{user}</span>
    </div>
    """, unsafe_allow_html=True)

    if st.button("Logout", use_container_width=True):
        for k in ["authenticated", "username", "api_key", "active_team", "active_project"]:
            del st.session_state[k]
        st.rerun()

    st.divider()

    # ── API KEY ──
    with st.expander("⚙️  Anthropic API Key", expanded=not st.session_state["api_key"]):
        key_input = st.text_input(
            "API Key", value=st.session_state["api_key"],
            type="password", placeholder="sk-ant-...",
            help="Stored only in this browser session.",
        )
        if key_input != st.session_state["api_key"]:
            st.session_state["api_key"] = key_input
        if st.session_state["api_key"]:
            st.success("✓ Claude AI mode active")
        else:
            st.warning("No key → Smart Fallback mode")

    st.divider()

    # ── TEAM ──
    st.markdown("##### 👥  Teams")
    teams = get_user_teams(user)

    with st.expander("＋  New Team"):
        new_team = st.text_input("Team name", key="new_team_name")
        if st.button("Create Team", key="btn_create_team", type="primary"):
            ok, msg = create_team(new_team, user)
            st.success(msg) if ok else st.error(msg)
            if ok: st.rerun()

    if teams:
        selected_team_name = st.selectbox("Active Team", [t["name"] for t in teams], key="sel_team")
        st.session_state["active_team"] = next(t for t in teams if t["name"] == selected_team_name)
    else:
        st.info("Create a team to get started.")
        st.session_state["active_team"] = None

    st.divider()

    # ── PROJECT ──
    st.markdown("##### 📁  Projects")
    if st.session_state["active_team"]:
        team_id  = st.session_state["active_team"]["id"]
        projects = get_team_projects(team_id)

        with st.expander("＋  New Project"):
            new_proj = st.text_input("Project name", key="new_proj_name")
            if st.button("Create Project", key="btn_create_proj", type="primary"):
                ok, msg = create_project(new_proj, team_id)
                st.success(msg) if ok else st.error(msg)
                if ok: st.rerun()

        if projects:
            sel_proj_name = st.selectbox("Active Project", [p["name"] for p in projects], key="sel_proj")
            st.session_state["active_project"] = next(p for p in projects if p["name"] == sel_proj_name)
        else:
            st.info("Create a project first.")
            st.session_state["active_project"] = None
    else:
        st.session_state["active_project"] = None

# ─────────────────────────────────────────────────────────────
#  MAIN AREA
# ─────────────────────────────────────────────────────────────
active_team    = st.session_state.get("active_team")
active_project = st.session_state.get("active_project")
team_label     = active_team["name"]    if active_team    else "—"
proj_label     = active_project["name"] if active_project else "—"

st.markdown(f"""
<div style='display:flex;align-items:center;justify-content:space-between;padding:8px 0 24px;'>
  <div>
    <h1 style='margin:0;font-size:2rem;letter-spacing:0.04em;'>
      🧬 AI Test Case Generator
    </h1>
    <span style='font-size:0.82rem;color:#64748b;'>{team_label}  ›  {proj_label}</span>
  </div>
  <div style='text-align:right;'>
    <div class='mode-badge {"mode-ai" if st.session_state["api_key"] else "mode-fallback"}'>
      {"⚡ CLAUDE AI MODE" if st.session_state["api_key"] else "🔁 SMART FALLBACK MODE"}
    </div>
    <div style='font-size:0.75rem;color:#64748b;margin-top:6px;font-family:"Space Mono",monospace;'>
      {datetime.now():%d %b %Y · %H:%M}
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

if not active_team or not active_project:
    st.info("👈  Select or create a **Team** and a **Project** in the sidebar to begin.")
    st.stop()

project_id = active_project["id"]

# ── Stats row ──
stats        = get_stats(project_id)
history_rows = get_project_history(project_id)
total_tc     = sum(
    len(json.loads(r["result"])) if r["result"].strip().startswith("[") else 15
    for r in history_rows
) if history_rows else 0

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Generation Runs",   stats["runs"])
c2.metric("Total Test Cases",  total_tc)
c3.metric("Contributors",      stats["contributors"])
c4.metric("Cases per Run",     "15")
c5.metric("Engine",            "Claude AI" if st.session_state["api_key"] else "Smart Fallback")

st.divider()

# ─────────────────────────────────────────────────────────────
#  TABS
# ─────────────────────────────────────────────────────────────
tab_gen, tab_hist = st.tabs(["🚀  Generate", "📜  History"])

# ══════════════════════════════════════════════════════════════
#  TAB 1 — GENERATE
# ══════════════════════════════════════════════════════════════
with tab_gen:
    st.markdown("#### Upload User Story Files")
    st.caption("Supported: `.txt` · `.pdf` · `.docx`  |  Multiple files accepted")

    uploaded_files = st.file_uploader(
        "Drop files here",
        type=["txt", "pdf", "docx"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

    with st.expander("✏️  Or enter a story manually"):
        manual_story = st.text_area(
            "User Story",
            placeholder="As a user, I want to reset my password via OTP so that I can regain access securely.",
            height=120,
            label_visibility="collapsed",
        )

    col_btn, col_info = st.columns([1, 3])
    with col_btn:
        generate_btn = st.button("🚀  Generate 15 Test Cases", type="primary", use_container_width=True)

    with col_info:
        if not st.session_state["api_key"]:
            st.warning("⚠️  No API key — Smart Dynamic Fallback will generate 15 context-aware cases.")
        else:
            st.success("✅  Claude AI will generate 15 deep, story-specific test cases.")

    if generate_btn:
        stories: list[str] = []

        for f in uploaded_files:
            text  = read_uploaded_file(f)
            parts = [s.strip() for s in re.split(r"\n{2,}", text) if s.strip()]
            stories.extend(parts)

        if manual_story.strip():
            stories.append(manual_story.strip())

        if not stories:
            st.error("Please upload at least one file or enter a story manually.")
        else:
            all_frames: list[pd.DataFrame] = []
            db = get_db()

            progress = st.progress(0, text="Initialising engine…")
            for i, story in enumerate(stories):
                progress.progress(
                    (i + 1) / len(stories),
                    text=f"Story {i+1}/{len(stories)}: {story[:70]}…"
                )
                df, engine_used = generate_test_cases(story, st.session_state["api_key"])
                df.insert(0, "User Story", story[:80])
                all_frames.append(df)

                db.execute(
                    "INSERT INTO testcases (username, project_id, story, result, engine) VALUES (?,?,?,?,?)",
                    (user, project_id, story,
                     df.drop(columns=["User Story"]).to_json(orient="records"),
                     engine_used),
                )
                db.commit()

            progress.empty()
            combined = pd.concat(all_frames, ignore_index=True)

            # ── Result banner ──
            badge_cls   = "mode-ai" if st.session_state["api_key"] else "mode-fallback"
            badge_label = "Claude AI" if st.session_state["api_key"] else "Smart Fallback"
            st.markdown(f"""
            <div style='background:#1a2236;border:1px solid #1f2d45;border-radius:12px;
                        padding:16px 20px;margin-bottom:16px;display:flex;
                        align-items:center;justify-content:space-between;'>
              <span style='font-family:"Space Mono",monospace;color:#00e5c3;font-size:1rem;'>
                ✅ Generated <strong>{len(combined)}</strong> test cases from
                <strong>{len(stories)}</strong> stor{'y' if len(stories)==1 else 'ies'}
              </span>
              <span class='mode-badge {badge_cls}'>Engine: {badge_label}</span>
            </div>
            """, unsafe_allow_html=True)

            # ── Styled dataframe ──
            def _style_priority(val):
                colors = {"High": "#ef4444", "Medium": "#f59e0b", "Low": "#22c55e"}
                c = colors.get(val, "transparent")
                return f"color:{c};font-weight:600;"

            def _style_type(val):
                colors = {
                    "Security":    "#a78bfa",
                    "Negative":    "#f87171",
                    "Performance": "#60a5fa",
                    "Boundary":    "#34d399",
                    "UX/UI":       "#f9a8d4",
                    "Functional":  "#00e5c3",
                }
                c = colors.get(val, "#e2e8f0")
                return f"color:{c};"

            try:
                styled = (combined.style
                          .map(_style_priority, subset=["Priority"])
                          .map(_style_type,     subset=["Type"])
                          .set_properties(**{"text-align": "left"}))
            except AttributeError:
                styled = (combined.style
                          .applymap(_style_priority, subset=["Priority"])
                          .applymap(_style_type,     subset=["Type"])
                          .set_properties(**{"text-align": "left"}))

            st.dataframe(styled, use_container_width=True, height=480)

            # ── Export ──
            st.markdown("#### 📤  Export")
            ex1, ex2 = st.columns(2)
            file_ext = "xlsx" if openpyxl else "csv"
            ex1.download_button(
                f"📥  Download {file_ext.upper()}",
                data=df_to_excel_bytes(combined),
                file_name=f"testcases_{proj_label}_{datetime.now():%Y%m%d_%H%M%S}.{file_ext}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                      if openpyxl else "text/csv",
                type="primary",
            )
            ex2.download_button(
                "📋  Download CSV",
                data=combined.to_csv(index=False).encode(),
                file_name=f"testcases_{proj_label}_{datetime.now():%Y%m%d_%H%M%S}.csv",
                mime="text/csv",
            )

            # ── Summary charts ──
            st.markdown("#### 📊  Summary")
            s1, s2, s3 = st.columns(3)
            with s1:
                st.markdown("**By Priority**")
                st.bar_chart(combined["Priority"].value_counts())
            with s2:
                st.markdown("**By Type**")
                st.bar_chart(combined["Type"].value_counts())
            with s3:
                st.markdown("**By Complexity**")
                st.bar_chart(combined["Complexity"].value_counts())

# ══════════════════════════════════════════════════════════════
#  TAB 2 — HISTORY
# ══════════════════════════════════════════════════════════════
with tab_hist:
    st.markdown("#### Recent Generations")

    if not history_rows:
        st.info("No test cases generated for this project yet.")
    else:
        for row in history_rows:
            ts      = row.get("created_at") or "—"
            story   = row["story"][:80] + ("…" if len(row["story"]) > 80 else "")
            engine  = row.get("engine") or "—"
            badge   = "mode-ai" if engine == "Claude AI" else "mode-fallback"

            with st.expander(f"🗂  {story}  ·  `{ts}`"):
                st.markdown(f"<span class='mode-badge {badge}'>{engine}</span>",
                            unsafe_allow_html=True)
                try:
                    cases = json.loads(row["result"])
                    df    = pd.DataFrame(cases)
                    st.dataframe(df, use_container_width=True)
                    st.download_button(
                        "📥  Download this batch",
                        data=df_to_excel_bytes(df),
                        file_name=f"batch_{row['id']}.{'xlsx' if openpyxl else 'csv'}",
                        key=f"dl_{row['id']}",
                    )
                except Exception:
                    st.write(row["result"])

# ─────────────────────────────────────────────────────────────
#  FOOTER
# ─────────────────────────────────────────────────────────────
st.divider()
st.markdown("""
<p style='text-align:center;color:#334155;font-size:0.75rem;
          font-family:"Space Mono",monospace;padding:8px 0 16px;'>
  QA·AI Ultimate Hybrid — 15 Cases · AI + Smart Fallback · Built with Streamlit & Anthropic Claude
</p>
""", unsafe_allow_html=True)
