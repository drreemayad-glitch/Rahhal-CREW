import os
import streamlit as st
from datetime import datetime
from docx import Document

from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT

try:
    from openai import OpenAI
except Exception:
    OpenAI = None


DEFAULT_MODEL = "gpt-4.1-mini"


# ---------------------------
# SYSTEM PROMPT
# ---------------------------

def load_prompt() -> str:
    try:
        with open("rahhal_prompt.txt", "r", encoding="utf-8") as f:
            return f.read()
    except Exception:
        return "Rahhal CREW system active. Ask one focused question at a time."


RAHHAL_SYSTEM = load_prompt()


# ---------------------------
# OPENAI CLIENT
# ---------------------------

def get_client():
    if OpenAI is None:
        st.sidebar.error("OpenAI library missing.")
        return None

    key = os.getenv("OPENAI_API_KEY")
    if not key:
        st.sidebar.error("Missing OPENAI_API_KEY in Secrets.")
        return None

    return OpenAI(api_key=key)


# ---------------------------
# FINAL PACKAGE FORMAT
# ---------------------------

def build_final_package_override():
    return """
For this response only:

Output ONLY markdown tables. No text outside tables.

System Mapping Summary:
Exercise Level | Format | Hazard | Systems | Stakeholders | Escalation Authority | Duration | Constraints

Objective Matrix:
Objective | Structural Dimension | Measurable Indicator | Linked Capability

Scenario Brief:
Field | Content

Inject Architecture:
Wave | Time | Inject Format | Linked Objective | Decision Focus | Discussion Questions | What to Observe

Facilitator Evaluation Matrix:
Objective | Structural Dimension | Expected Actions | Performance Indicators | Evidence Source | Observer Notes

Coverage Validation Summary:
Item | Status | Notes

Keep cell text concise.
Do not use the pipe character inside cells.
"""


# ---------------------------
# SESSION
# ---------------------------

def ensure_session():
    if "messages" not in st.session_state:
        st.session_state.messages = [
            {"role": "system", "content": RAHHAL_SYSTEM},
            {
                "role": "assistant",
                "content": (
                    "Rahhal CREW activated.\n\n"
                    "Select exercise format:\n"
                    "Discussion Based Tabletop or Functional Exercise?"
                ),
            },
        ]

    if "pending_input" not in st.session_state:
        st.session_state.pending_input = None

    if "temperature" not in st.session_state:
        st.session_state.temperature = 0.1


# ---------------------------
# MARKDOWN TABLE PARSER
# ---------------------------

def is_table_line(line):
    return line.strip().startswith("|") and line.strip().endswith("|")


def split_row(line):
    return [c.strip() for c in line.strip().strip("|").split("|")]


def style_cell(cell, bold=False):
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    for p in cell.paragraphs:
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for run in p.runs:
            run.font.size = Pt(9)
            run.bold = bold


def add_table(doc, header_line, rows):
    header = split_row(header_line)
    ncols = len(header)

    table = doc.add_table(rows=1, cols=ncols)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = True

    # Header
    for i, col in enumerate(header):
        cell = table.rows[0].cells[i]
        cell.text = col
        style_cell(cell, bold=True)

    # Rows
    for row_line in rows:
        row_data = split_row(row_line)
        row = table.add_row().cells
        for i in range(ncols):
            row[i].text = row_data[i] if i < len(row_data) else ""
            style_cell(row[i])


def add_markdown(doc, text):
    lines = text.splitlines()
    i = 0

    while i < len(lines):
        line = lines[i]

        if is_table_line(line):
            header = line
            i += 2  # skip separator
            rows = []

            while i < len(lines) and is_table_line(lines[i]):
                rows.append(lines[i])
                i += 1

            add_table(doc, header, rows)
            doc.add_paragraph("")
            continue

        if line.strip():
            doc.add_paragraph(line.strip())

        i += 1


# ---------------------------
# EXPORT WORD
# ---------------------------

def export_docx(messages):
    doc = Document()

    doc.add_paragraph(
        f"Generated: {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}"
    )
    doc.add_paragraph("")

    for m in messages:
        if m["role"] == "system":
            continue

        role = m["role"].capitalize()
        doc.add_paragraph(role).runs[0].bold = True
        add_markdown(doc, m["content"])
        doc.add_paragraph("")

    path = "rahhal_output.docx"
    doc.save(path)
    return path


# ---------------------------
# MODEL CALL
# ---------------------------

def run_model(user_input):
    st.session_state.messages.append({"role": "user", "content": user_input})

    if user_input.upper() == "FINAL PACKAGE":
        st.session_state.messages.append(
            {"role": "system", "content": build_final_package_override()}
        )

    client = get_client()
    if client is None:
        return

    response = client.chat.completions.create(
        model=DEFAULT_MODEL,
        temperature=st.session_state.temperature,
        messages=st.session_state.messages,
    )

    reply = response.choices[0].message.content
    st.session_state.messages.append({"role": "assistant", "content": reply})


# ---------------------------
# UI
# ---------------------------

st.set_page_config(page_title="Rahhal CREW", page_icon="🛡", layout="wide")
ensure_session()

st.title("Rahhal CREW")
st.caption("AI Augmented Crisis Readiness Exercise Design System")
st.divider()

with st.sidebar:
    st.header("Control Panel")

    st.session_state.temperature = st.slider(
        "Creativity", 0.0, 1.0, st.session_state.temperature, 0.05
    )

    if st.button("Reset"):
        del st.session_state["messages"]
        st.rerun()

    if st.button("Generate structured package"):
        st.session_state.pending_input = "FINAL PACKAGE"
        st.rerun()

    if st.button("Export Word"):
        msgs = [m for m in st.session_state.messages if m["role"] != "system"]
        file_path = export_docx(msgs)
        with open(file_path, "rb") as f:
            st.download_button(
                "Download Word File",
                f,
                file_name=file_path,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

tab_chat, tab_help = st.tabs(["Chat", "How it works"])

with tab_help:
    st.markdown("""
• Select exercise format  
• Answer Rahhal step by step  
• Click Generate structured package  
• Export Word file  
""")

with tab_chat:
    for m in st.session_state.messages:
        avatar = "🛡" if m["role"] == "assistant" else "👤"
        with st.chat_message(m["role"], avatar=avatar):
            st.write(m["content"])

    user_input = st.chat_input("Write your answer here")

    if user_input:
        run_model(user_input)
        st.rerun()
