# ============================================================
#  QA·AI  —  Professional AI-Powered Test Case Generator
#  Streamlit · SQLite · Anthropic Claude API
# ============================================================

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

# File readers
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
    import openpyxl                         # noqa: F401 – needed for pd.to_excel
except ImportError:
    openpyxl = None

# ─────────────────────────────────────────────────────────────
#  PAGE CONFIG  (must be first Streamlit call)
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="QA·AI — Test Case Generator",
    page_icon="🧪",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────
#  GLOBAL CSS  — dark-glass design, electric-teal accents
# ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&family=DM+Sans:wght@300;400;500;600;700&display=swap');

/* ── reset & base ── */
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
    --text:      #e2e8f0;
    --muted:     #64748b;
    --radius:    12px;
}

html, body, [data-testid="stAppViewContainer"] {
    background: var(--bg) !important;
    color: var(--text) !important;
    font-family: 'DM Sans', sans-serif !important;
}

/* ── sidebar ── */
[data-testid="stSidebar"] {
    background: var(--surface) !important;
    border-right: 1px solid var(--border) !important;
}
[data-testid="stSidebar"] * { color: var(--text) !important; }

/* ── headings ── */
h1, h2, h3 {
    font-family: 'Space Mono', monospace !important;
    color: var(--text) !important;
}

/* ── metric cards ── */
[data-testid="stMetric"] {
    background: var(--surface2) !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--radius) !important;
    padding: 16px 20px !important;
}
[data-testid="stMetricLabel"] { color: var(--muted) !important; font-size: 0.78rem !important; }
[data-testid="stMetricValue"] { color: var(--teal) !important; font-family: 'Space Mono', monospace !important; }

/* ── inputs ── */
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

/* ── primary button ── */
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

/* ── secondary buttons ── */
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

/* ── dataframe ── */
[data-testid="stDataFrame"] {
    border: 1px solid var(--border) !important;
    border-radius: var(--radius) !important;
    overflow: hidden !important;
}

/* ── success / warning / error ── */
[data-testid="stAlert"] {
    border-radius: var(--radius) !important;
    border: 1px solid var(--border) !important;
}

/* ── tabs ── */
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

/* ── divider ── */
hr { border-color: var(--border) !important; }

/* ── spinner ── */
[data-testid="stSpinner"] { color: var(--teal) !important; }

/* ── expander ── */
[data-testid="stExpander"] {
    background: var(--surface2) !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--radius) !important;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
#  DATABASE  (per-session connection avoids threading errors)
# ─────────────────────────────────────────────────────────────
DB_PATH = Path("qa_ai.db")

def get_db():
    """Return a per-session SQLite connection stored in session_state."""
    if "_db" not in st.session_state:
        conn = sqlite3.connect(str(DB_PATH), check_same_thread=False)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        _create_schema(conn)
        st.session_state["_db"] = conn
    return st.session_state["_db"]

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
        created_at TEXT DEFAULT (datetime('now'))
    );
    """)
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
    """Extract text from .txt / .pdf / .docx uploads."""
    name = file.name.lower()
    try:
        if name.endswith(".txt"):
            return file.read().decode("utf-8", errors="replace")

        elif name.endswith(".pdf"):
            if PyPDF2 is None:
                return "[PyPDF2 not installed — cannot read PDF]"
            reader = PyPDF2.PdfReader(file)
            pages = [p.extract_text() or "" for p in reader.pages]
            return "\n".join(pages)

        elif name.endswith(".docx"):
            if DocxDocument is None:
                return "[python-docx not installed — cannot read DOCX]"
            doc = DocxDocument(file)
            return "\n".join(p.text for p in doc.paragraphs)

    except Exception as exc:
        return f"[Error reading file: {exc}]"
    return ""

# ─────────────────────────────────────────────────────────────
#  AI — REAL ANTHROPIC CALL
# ─────────────────────────────────────────────────────────────
_SYSTEM_PROMPT = """You are an expert QA engineer. Given a user story, generate exactly 5 structured test cases.

Respond ONLY with a valid JSON array — no markdown, no prose, no code fences.
Each object must have these keys:
  "id"              : "TC_1" .. "TC_5"
  "scenario"        : one-sentence test scenario
  "preconditions"   : what must be true before the test
  "steps"           : numbered steps as a single string
  "expected_result" : what the system should do
  "priority"        : one of High | Medium | Low
  "complexity"      : one of High | Medium | Low
  "type"            : one of Functional | Negative | Boundary | Performance | Security | UX

Return ONLY the JSON array."""

def _fallback_cases(story: str) -> list[dict]:
    """Deterministic fallback used when AI is unavailable."""
    types  = ["Functional", "Negative", "Boundary", "UX", "Security"]
    prios  = ["High", "High", "Medium", "Medium", "Low"]
    return [
        {
            "id": f"TC_{i}",
            "scenario":         f"Verify scenario {i} for: {story[:60]}",
            "preconditions":    "User is logged in and the feature is enabled.",
            "steps":            f"1. Navigate to feature\n2. Perform action {i}\n3. Observe outcome",
            "expected_result":  "System responds correctly without errors.",
            "priority":         prios[i-1],
            "complexity":       "Medium",
            "type":             types[i-1],
        }
        for i in range(1, 6)
    ]

def generate_test_cases(story: str, api_key: str) -> pd.DataFrame:
    """Call Claude; fall back gracefully on any error."""
    cases = None

    if anthropic and api_key:
        try:
            client = anthropic.Anthropic(api_key=api_key)
            message = client.messages.create(
                model="claude-opus-4-6",
                max_tokens=2048,
                system=_SYSTEM_PROMPT,
                messages=[{"role": "user", "content": story}],
            )
            raw = message.content[0].text.strip()
            # Strip accidental code fences
            raw = re.sub(r"^```[a-z]*\n?|```$", "", raw, flags=re.MULTILINE).strip()
            cases = json.loads(raw)
        except json.JSONDecodeError as e:
            st.warning(f"AI returned invalid JSON — using fallback. ({e})")
        except Exception as e:
            st.warning(f"AI call failed — using fallback. ({e})")

    if not cases:
        cases = _fallback_cases(story)

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
    return df

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
        FROM teams t
        JOIN team_members tm ON t.id = tm.team_id
        WHERE tm.username = ?
        ORDER BY t.id DESC
    """, (username,)).fetchall()
    return [dict(r) for r in rows]

def create_project(name: str, team_id: int) -> tuple[bool, str]:
    if not name.strip():
        return False, "Project name cannot be empty."
    db = get_db()
    db.execute("INSERT INTO projects (name, team_id) VALUES (?,?)",
               (name.strip(), team_id))
    db.commit()
    return True, f"Project '{name}' created."

def get_team_projects(team_id: int) -> list[dict]:
    db = get_db()
    rows = db.execute("""
        SELECT id, name, created_at FROM projects
        WHERE team_id = ? ORDER BY id DESC
    """, (team_id,)).fetchall()
    return [dict(r) for r in rows]

def get_project_history(project_id: int, limit: int = 20) -> list[dict]:
    db = get_db()
    rows = db.execute("""
        SELECT id, username, story, result, created_at
        FROM testcases WHERE project_id = ?
        ORDER BY id DESC LIMIT ?
    """, (project_id, limit)).fetchall()
    return [dict(r) for r in rows]

def get_stats(project_id: int) -> dict:
    db = get_db()
    r = db.execute("""
        SELECT COUNT(*) as runs,
               COUNT(DISTINCT username) as contributors
        FROM testcases WHERE project_id = ?
    """, (project_id,)).fetchone()
    return {"runs": r["runs"], "contributors": r["contributors"]}

# ─────────────────────────────────────────────────────────────
#  EXCEL EXPORT
# ─────────────────────────────────────────────────────────────
def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    if openpyxl is None:
        # Fallback to CSV
        return df.to_csv(index=False).encode()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Test Cases")
        ws = writer.sheets["Test Cases"]
        # Auto-width columns
        for col in ws.columns:
            max_len = max((len(str(cell.value or "")) for cell in col), default=10)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 60)
    return buf.getvalue()

# ─────────────────────────────────────────────────────────────
#  SESSION STATE BOOTSTRAP
# ─────────────────────────────────────────────────────────────
for _key, _default in [
    ("authenticated", False),
    ("username",      ""),
    ("api_key",       ""),
    ("active_team",   None),
    ("active_project",None),
]:
    if _key not in st.session_state:
        st.session_state[_key] = _default

# ─────────────────────────────────────────────────────────────
#  SIDEBAR — AUTH + NAV
# ─────────────────────────────────────────────────────────────
with st.sidebar:
    # ── Logo ──
    st.markdown("""
    <div style='padding:20px 0 8px;'>
      <span style='font-family:"Space Mono",monospace;font-size:1.5rem;
                   color:#00e5c3;letter-spacing:0.08em;'>QA·AI</span>
      <br>
      <span style='font-size:0.72rem;color:#64748b;letter-spacing:0.15em;
                   text-transform:uppercase;'>Test Case Generator</span>
    </div>
    <hr style='margin:8px 0 16px;'>
    """, unsafe_allow_html=True)

    # ── AUTH ──
    if not st.session_state["authenticated"]:
        auth_tab, reg_tab = st.tabs(["🔑  Login", "✨  Register"])

        with auth_tab:
            login_user_input = st.text_input("Username", key="li_user")
            login_pwd_input  = st.text_input("Password", type="password", key="li_pwd")
            if st.button("Login", key="btn_login", use_container_width=True, type="primary"):
                if login_user(login_user_input, login_pwd_input):
                    st.session_state["authenticated"] = True
                    st.session_state["username"]       = login_user_input.strip()
                    st.rerun()
                else:
                    st.error("Invalid credentials.")

        with reg_tab:
            reg_user_input = st.text_input("Username", key="reg_user")
            reg_pwd_input  = st.text_input("Password", type="password", key="reg_pwd")
            reg_pwd2_input = st.text_input("Confirm Password", type="password", key="reg_pwd2")
            if st.button("Create Account", key="btn_reg", use_container_width=True, type="primary"):
                if reg_pwd_input != reg_pwd2_input:
                    st.error("Passwords do not match.")
                else:
                    ok, msg = register_user(reg_user_input, reg_pwd_input)
                    if ok:
                        st.success(msg)
                    else:
                        st.error(msg)
        st.stop()

    # ── LOGGED-IN NAV ──
    st.markdown(f"""
    <div style='display:flex;align-items:center;gap:10px;padding:4px 0 12px;'>
      <div style='width:34px;height:34px;border-radius:50%;background:#00e5c3;
                  display:flex;align-items:center;justify-content:center;
                  font-family:"Space Mono",monospace;font-weight:700;color:#0b0f1a;'>
        {st.session_state["username"][0].upper()}
      </div>
      <span style='font-size:0.9rem;'>{st.session_state["username"]}</span>
    </div>
    """, unsafe_allow_html=True)

    if st.button("Logout", use_container_width=True):
        for k in ["authenticated", "username", "api_key",
                  "active_team", "active_project"]:
            del st.session_state[k]
        st.rerun()

    st.divider()

    # ── API KEY ──
    with st.expander("⚙️  Anthropic API Key", expanded=not st.session_state["api_key"]):
        key_input = st.text_input(
            "API Key",
            value=st.session_state["api_key"],
            type="password",
            placeholder="sk-ant-...",
            help="Your key is stored only in this browser session.",
        )
        if key_input != st.session_state["api_key"]:
            st.session_state["api_key"] = key_input
        if st.session_state["api_key"]:
            st.success("API key set ✓")
        else:
            st.warning("Add your API key to use Claude AI.")

    st.divider()

    # ── TEAM ──
    st.markdown("##### 👥  Teams")
    user = st.session_state["username"]
    teams = get_user_teams(user)
    team_names = [t["name"] for t in teams]

    with st.expander("＋  New Team"):
        new_team = st.text_input("Team name", key="new_team_name")
        if st.button("Create Team", key="btn_create_team", type="primary"):
            ok, msg = create_team(new_team, user)
            if ok:
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)

    if team_names:
        selected_team_name = st.selectbox("Active Team", team_names, key="sel_team")
        active_team = next(t for t in teams if t["name"] == selected_team_name)
        st.session_state["active_team"] = active_team
    else:
        st.info("Create a team to get started.")
        st.session_state["active_team"] = None

    st.divider()

    # ── PROJECT ──
    st.markdown("##### 📁  Projects")
    if st.session_state["active_team"]:
        team_id  = st.session_state["active_team"]["id"]
        projects = get_team_projects(team_id)
        proj_names = [p["name"] for p in projects]

        with st.expander("＋  New Project"):
            new_proj = st.text_input("Project name", key="new_proj_name")
            if st.button("Create Project", key="btn_create_proj", type="primary"):
                ok, msg = create_project(new_proj, team_id)
                if ok:
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)

        if proj_names:
            sel_proj_name = st.selectbox("Active Project", proj_names, key="sel_proj")
            active_proj = next(p for p in projects if p["name"] == sel_proj_name)
            st.session_state["active_project"] = active_proj
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

# ── Header ──
team_label = active_team["name"]    if active_team    else "—"
proj_label = active_project["name"] if active_project else "—"

st.markdown(f"""
<div style='display:flex;align-items:center;justify-content:space-between;
            padding:8px 0 24px;'>
  <div>
    <h1 style='margin:0;font-size:2rem;letter-spacing:0.04em;'>
      🧪 AI Test Case Generator
    </h1>
    <span style='font-size:0.82rem;color:#64748b;'>
      {team_label}  ›  {proj_label}
    </span>
  </div>
  <div style='font-size:0.78rem;color:#00e5c3;font-family:"Space Mono",monospace;
              text-align:right;'>
    {datetime.now():%d %b %Y · %H:%M}
  </div>
</div>
""", unsafe_allow_html=True)

# Guard: need team + project
if not active_team or not active_project:
    st.info("👈  Select or create a **Team** and a **Project** in the sidebar to begin.")
    st.stop()

project_id = active_project["id"]

# ── Stats row ──
stats = get_stats(project_id)
history_rows = get_project_history(project_id)
total_tc = sum(
    len(json.loads(r["result"])) if r["result"].startswith("[") else 5
    for r in history_rows
) if history_rows else 0

c1, c2, c3, c4 = st.columns(4)
c1.metric("Generation Runs",   stats["runs"])
c2.metric("Total Test Cases",  total_tc)
c3.metric("Contributors",      stats["contributors"])
c4.metric("API Engine",        "Claude AI" if st.session_state["api_key"] else "Fallback")

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
    st.caption("Supported formats: `.txt`, `.pdf`, `.docx`  |  Multiple files accepted.")

    uploaded_files = st.file_uploader(
        "Drop files here",
        type=["txt", "pdf", "docx"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

    # Optional manual story input
    with st.expander("✏️  Or enter a story manually"):
        manual_story = st.text_area(
            "User Story",
            placeholder="As a user, I want to … so that …",
            height=110,
            label_visibility="collapsed",
        )

    col_btn, col_info = st.columns([1, 3])
    with col_btn:
        generate_btn = st.button("🚀  Generate Test Cases", type="primary",
                                  use_container_width=True)

    if not st.session_state["api_key"]:
        col_info.warning("⚠️  No API key — fallback (deterministic) test cases will be used.")

    if generate_btn:
        stories: list[str] = []

        # Collect from files
        for f in uploaded_files:
            text = read_uploaded_file(f)
            parts = [s.strip() for s in re.split(r"\n{2,}", text) if s.strip()]
            stories.extend(parts)

        # Collect manual
        if manual_story.strip():
            stories.append(manual_story.strip())

        if not stories:
            st.error("Please upload at least one file or enter a story manually.")
        else:
            all_frames: list[pd.DataFrame] = []
            db = get_db()

            progress = st.progress(0, text="Generating…")
            for i, story in enumerate(stories):
                progress.progress((i + 1) / len(stories),
                                  text=f"Story {i+1}/{len(stories)}: {story[:60]}…")

                df = generate_test_cases(story, st.session_state["api_key"])

                # Tag with source story
                df.insert(0, "User Story", story[:80])
                all_frames.append(df)

                # Persist
                db.execute(
                    "INSERT INTO testcases (username, project_id, story, result) VALUES (?,?,?,?)",
                    (user, project_id, story, df.drop(columns=["User Story"]).to_json(orient="records")),
                )
                db.commit()

            progress.empty()
            combined = pd.concat(all_frames, ignore_index=True)

            st.success(f"✅  Generated **{len(combined)}** test cases from "
                       f"**{len(stories)}** story/stories.")

            # ── Priority colour map ──
            def _style_priority(val):
                colors = {"High": "#ef4444", "Medium": "#f59e0b", "Low": "#22c55e"}
                c = colors.get(val, "transparent")
                return f"color:{c};font-weight:600;"

            styled = combined.style.applymap(
                _style_priority, subset=["Priority"]
            ).set_properties(**{"text-align": "left"})

            st.dataframe(styled, use_container_width=True, height=420)

            # ── Exports ──
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

            # ── Summary chart ──
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
            ts    = row["created_at"] or "—"
            story = row["story"][:80] + ("…" if len(row["story"]) > 80 else "")

            with st.expander(f"🗂  {story}  ·  `{ts}`"):
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
  QA·AI — Powered by Anthropic Claude &nbsp;·&nbsp; Built with Streamlit
</p>
""", unsafe_allow_html=True)
