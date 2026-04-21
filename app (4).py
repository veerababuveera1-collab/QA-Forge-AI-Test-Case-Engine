import streamlit as st

import pandas as pd

from crewai import Agent, Task, Crew

from langchain_openai import ChatOpenAI

from docx import Document

import uuid

 

# ---------- CONFIG ----------

st.set_page_config(page_title="AI Test Case Generator", layout="wide")

st.title("🧠 AI Test Case Generator – Enterprise Format")

st.caption("Upload a User Story → Get Excel-ready test cases with CT IDs")

 

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

    backstory="Telecom B2C expert (PLP, PDP, CMS, Financing, Credit law)",

    llm=llm

)

 

formatter_agent = Agent(

    role="Test Case Formatter",

    goal="Convert test cases into a fixed Excel-ready structure",

    llm=llm

)

 

# ---------- BUTTON ----------

if st.button("🚀 Generate Test Cases") and uploaded_file:

 

    user_story_text = read_file(uploaded_file)

 

    task_generate = Task(

        description=f"""

From the following user story, generate test cases.

 

STRICT RULES:

- Use ONLY these columns:

CT Test Case IDs | OSD Number | ACC FLOW ID | ACC | Test Case Description |

Test case steps | Expected results | Pre-condition | Post condition |

Dev Status | Testing Status | Comments | ETA of Dev | ETA of CT | TESTER

 

- Generate realistic telecom ecommerce test cases

- Dev Status = "Not Started"

- Testing Status = "Not Started"

- ETA fields = "TBD"

- TESTER = ""

- One row = One test case

 

USER STORY:

{user_story_text}

""",

        agent=qa_agent

    )

 

    task_format = Task(

        description="Format output strictly as table rows (pipe separated, Excel compatible).",

        agent=formatter_agent

    )

 

    crew = Crew(

        agents=[qa_agent, formatter_agent],

        tasks=[task_generate, task_format]

    )

 

    result = crew.kickoff()

 

    # ---------- PARSE OUTPUT ----------

    rows = []

    for line in result.split("\n"):

        if "|" in line and "CT Test Case" not in line:

            cols = [c.strip() for c in line.split("|")]

            if len(cols) == 15:

                rows.append(cols)

 

    columns = [

        "CT Test Case IDs", "OSD Number", "ACC FLOW ID", "ACC",

        "Test Case Description", "Test case steps", "Expected results",

        "Pre-condition", "Post condition", "Dev Status", "Testing Status",

        "Comments", "ETA of Dev", "ETA of CT", "TESTER"

    ]

 

    df = pd.DataFrame(rows, columns=columns)

 

    st.subheader("✅ Generated Test Cases")

    st.dataframe(df, use_container_width=True)

 

    # ---------- DOWNLOAD ----------

    st.download_button(

        "📥 Download as Excel",

        data=df.to_excel(index=False),

        file_name="AI_Generated_Test_Cases.xlsx"

    )

