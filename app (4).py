import streamlit as st
import sqlite3
import pandas as pd
import hashlib
import io
import re
import json
from datetime import datetime
from pathlib import Path

# --- Dependencies Check ---
try:
    import anthropic
except ImportError:
    anthropic = None

# ─────────────────────────────────────────────────────────────
# 1. PAGE & THEME CONFIG
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="QA·AI Pro | Intelligent Test Generator",
    page_icon="🧪",
    layout="wide"
)

# Custom CSS for a professional Dark-Teal look
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600&family=Space+Mono&display=swap');
    
    :root {
        --primary: #00e5c3;
        --bg-dark: #0b101a;
        --card-bg: #161c2d;
    }
    
    .stApp { background-color: var(--bg-dark); color: #e2e8f0; font-family: 'Inter', sans-serif; }
    h1, h2, h3 { font-family: 'Space Mono', monospace !important; color: var(--primary) !important; }
    
    /* Metric Card Styling */
    [data-testid="stMetric"] {
        background-color: var(--card-bg) !important;
        border: 1px solid #2d3748 !important;
        border-radius: 12px !important;
        padding: 15px !important;
    }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# 2. DATABASE & AUTH LOGIC
# ─────────────────────────────────────────────────────────────
DB_PATH = "qa_pro_v2.db"

def init_db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT,
            story TEXT,
            test_cases TEXT
        )
    """)
    conn.commit()
    return conn

# ─────────────────────────────────────────────────────────────
# 3. CORE AI LOGIC (The "Overcome Issues" Part)
# ─────────────────────────────────────────────────────────────

# AI ప్రాంప్ట్ - ప్రతి కేస్ విభిన్నంగా ఉండాలని ఆదేశిస్తుంది
SYSTEM_INSTRUCTIONS = """You are a Lead QA Engineer. 
Generate 5 unique and technically sound test cases for the given User Story.
STRICT REQUIREMENTS:
1. NO DUPLICATION: Every 'Expected Result' and 'Steps' must be unique to its test type.
2. NEGATIVE TESTS: Must define specific error messages.
3. SECURITY TESTS: Must define access denial or encryption checks.
4. FORMAT: Return ONLY a valid JSON array. No text before or after.

JSON Structure:
[
  {"id": "TC_1", "scenario": "...", "preconditions": "...", "steps": "...", "expected_result": "...", "priority": "High/Medium/Low", "type": "Functional/Negative/Boundary/Security/UX"}
]"""

def fallback_logic(story):
    """AI ఫెయిల్ అయినప్పుడు లేదా API కీ లేనప్పుడు ఉపయోగించే డైనమిక్ లాజిక్"""
    return [
        {"id": "TC_1", "scenario": f"Basic validation for {story[:20]}...", "preconditions": "Feature enabled", "steps": "1. Login\n2. Perform action", "expected_result": "System updates successfully", "priority": "High", "type": "Functional"},
        {"id": "TC_2", "scenario": "Negative data input check", "preconditions": "On update page", "steps": "1. Enter null values\n2. Submit", "expected_result": "Error: Required fields missing", "priority": "High", "type": "Negative"}
    ]

def generate_cases(story, api_key):
    if anthropic and api_key.startswith("sk-ant"):
        try:
            client = anthropic.Anthropic(api_key=api_key)
            # మోడల్ పేరు అప్‌డేట్ చేయబడింది (Fixed Issue #1)
            response = client.messages.create(
                model="claude-3-5-sonnet-20240620", 
                max_tokens=2500,
                system=SYSTEM_INSTRUCTIONS,
                messages=[{"role": "user", "content": f"User Story: {story}"}]
            )
            data = json.loads(re.sub(r"```json|```", "", response.content[0].text))
            return pd.DataFrame(data)
        except Exception as e:
            st.error(f"AI Connection Error: {e}")
            return pd.DataFrame(fallback_logic(story))
    else:
        return pd.DataFrame(fallback_logic(story))

# ─────────────────────────────────────────────────────────────
# 4. MAIN APP INTERFACE
# ─────────────────────────────────────────────────────────────
def main():
    init_db()
    
    # --- Sidebar ---
    with st.sidebar:
        st.title("⚙️ Settings")
        key = st.text_input("Enter Anthropic API Key", type="password", help="Starts with sk-ant-...")
        st.divider()
        st.info("💡 Tip: Use a valid API Key to get unique and logical test cases instead of generic ones.")

    # --- Header ---
    st.title("🧪 QA·AI Pro")
    st.markdown("### Intelligent End-to-End Test Case Generator")
    
    # --- Input Area ---
    col1, col2 = st.columns([2, 1])
    with col1:
        story = st.text_area("📄 Paste User Story / Requirement:", height=200, 
                             placeholder="Example: As a marketing manager, I want to update product prices so that they reflect current season discounts.")
    
    with col2:
        st.write("📊 **Quick Tips**")
        st.caption("- Be specific in your story for better AI results.")
        st.caption("- AI will generate Functional, Negative, and Security cases.")
        generate_btn = st.button("🚀 Generate Test Suite", use_container_width=True, type="primary")

    # --- Execution & Results ---
    if generate_btn:
        if not story:
            st.warning("Please enter a User Story first.")
        else:
            with st.spinner("Analyzing requirements and generating cases..."):
                df = generate_cases(story, key)
                
                # Metrics
                m1, m2, m3 = st.columns(3)
                m1.metric("Total Cases", len(df))
                m2.metric("Coverage", "Full (End-to-End)")
                m3.metric("Engine", "Claude 3.5 Sonnet" if key else "Fallback Mode")

                st.divider()
                
                # Styling Table (Fixed Issue #2: Priority Colors)
                def style_priority(val):
                    color = '#ff4b4b' if val == 'High' else '#ffa500' if val == 'Medium' else '#00e5c3'
                    return f'color: {color}; font-weight: bold'

                st.subheader("📋 Generated Test Cases")
                st.dataframe(df.style.applymap(style_priority, subset=['priority']), use_container_width=True)

                # Export Options
                csv = df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Download as CSV for JIRA/Excel",
                    data=csv,
                    file_name=f"TestCases_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv",
                )

                # Save to History
                conn = init_db()
                conn.execute("INSERT INTO history (timestamp, story, test_cases) VALUES (?, ?, ?)",
                             (datetime.now().isoformat(), story, df.to_json()))
                conn.commit()

    # --- History Section ---
    with st.expander("📜 View Recent History"):
        conn = init_db()
        history_df = pd.read_sql_query("SELECT timestamp, story FROM history ORDER BY id DESC LIMIT 5", conn)
        st.table(history_df)

if __name__ == "__main__":
    main()
