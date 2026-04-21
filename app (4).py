__import__('pysqlite3')
import sys
sys.modules['sqlite3'] = sys.modules.pop('pysqlite3')
import streamlit as st
import pandas as pd
from crewai import Agent, Task, Crew
from langchain_openai import ChatOpenAI
from docx import Document
import io

# ---------- CONFIG ----------
st.set_page_config(page_title="AI Test Case Generator", layout="wide")
st.title("🧠 AI Test Case Generator – Enterprise Format")
st.caption("Upload a User Story → Get Excel-ready test cases with CT IDs")

# Ensure you have your OPENAI_API_KEY set in your environment variables
llm = ChatOpenAI(model="gpt-4o", temperature=0.2)

# ---------- FILE UPLOAD ----------
uploaded_file = st.file_uploader(
    "📤 Upload User Story (Word or Text)",
    type=["docx", "txt"]
)

def read_file(file):
    if file.name.endswith(".docx"):
        doc = Document(file)
        return "\n".join([p.text for p in doc.paragraphs])
    return file.read().decode("utf-8")

# ---------- AGENTS ----------
qa_agent = Agent(
    role="Senior QA Engineer",
    goal="Generate structured enterprise test cases",
    backstory="Telecom B2C expert (PLP, PDP, CMS, Financing, Credit law). You excel at detail.",
    llm=llm,
    allow_delegation=False
)

formatter_agent = Agent(
    role="Test Case Formatter",
    goal="Convert test cases into a clean Markdown table format",
    backstory="You are a data extraction specialist who ensures tables are perfectly formatted for CSV/Excel parsing.",
    llm=llm,
    allow_delegation=False
)

# ---------- PROCESSING LOGIC ----------
if st.button("🚀 Generate Test Cases") and uploaded_file:
    user_story_text = read_file(uploaded_file)

    with st.spinner("CrewAI is working on your test cases..."):
        task_generate = Task(
            description=f"""
            From the following user story, generate 5-10 detailed test cases.
            
            STRICT COLUMNS:
            CT Test Case IDs | OSD Number | ACC FLOW ID | ACC | Test Case Description | Test case steps | Expected results | Pre-condition | Post condition | Dev Status | Testing Status | Comments | ETA of Dev | ETA of CT | TESTER
            
            USER STORY:
            {user_story_text}
            """,
            expected_output="A list of detailed test cases following the telecom B2C domain standards.",
            agent=qa_agent
        )

        task_format = Task(
            description="""Format the previous test cases into a SINGLE Markdown table. 
            Ensure every row has exactly 15 columns. Use 'TBD' for empty date fields and 'Not Started' for status fields.
            Do not include any introductory text, only the table.""",
            expected_output="A single markdown table with 15 columns.",
            agent=formatter_agent
        )

        crew = Crew(
            agents=[qa_agent, formatter_agent],
            tasks=[task_generate, task_format]
        )

        # crew.kickoff() returns a CrewOutput object; use .raw to get the string
        result = crew.kickoff().raw

        # ---------- PARSE OUTPUT ----------
        rows = []
        for line in result.split("\n"):
            if "|" in line and "---" not in line:
                # Split and clean columns, removing leading/trailing pipes
                cols = [c.strip() for c in line.split("|") if c.strip() != ""]
                if len(cols) == 15 and "CT Test Case" not in cols[0]:
                    rows.append(cols)

        columns = [
            "CT Test Case IDs", "OSD Number", "ACC FLOW ID", "ACC",
            "Test Case Description", "Test case steps", "Expected results",
            "Pre-condition", "Post condition", "Dev Status", "Testing Status",
            "Comments", "ETA of Dev", "ETA of CT", "TESTER"
        ]

        if rows:
            df = pd.DataFrame(rows, columns=columns)
            st.subheader("✅ Generated Test Cases")
            st.dataframe(df, use_container_width=True)

            # ---------- DOWNLOAD (The Fix) ----------
            # Pandas .to_excel requires a buffer for Streamlit download buttons
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='TestCases')
            
            st.download_button(
                label="📥 Download as Excel",
                data=buffer.getvalue(),
                file_name="AI_Generated_Test_Cases.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Could not parse the AI output into a table. Please try again.")
            st.text(result) # Show raw output for debugging
