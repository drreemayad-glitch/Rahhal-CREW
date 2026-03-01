import os
import streamlit as st
from datetime import datetime
from docx import Document

try:
    from openai import OpenAI
except Exception:
    OpenAI = None


def load_prompt() -> str:
    try:
        with open("rahhal_prompt.txt", "r", encoding="utf-8") as f:
            return f.read()
    except Exception:
        return "You are Rahhal CREW. Ask one focused question at a time."


RAHHAL_SYSTEM = load_prompt()
DEFAULT_MODEL = "gpt-4.1-mini"


def get_client_or_sidebar_error():
    if OpenAI is None:
        st.sidebar.error("OpenAI library not installed. Check requirements.txt.")
        return None

    key = (os.getenv("OPENAI_API_KEY") or "").strip()
    if not key:
        st.sidebar.error("Missing OPENAI_API_KEY. Add it in Streamlit Cloud Secrets.")
        return None

    return OpenAI(api_key=key)


def _is_md_table_line(line: str) -> bool:
    s = line.strip()
    return s.startswith("|") and s.endswith("|") and s.count("|") >= 3


def _is_md_separator_line(line: str) -> bool:
    s = line.strip().replace(" ", "")
    if not (s.startswith("|") and s.endswith("|")):
        return False
    cells = [c for c in s.strip("|").split("|")]
    for c in cells:
        if c and not all(ch in "-:" for ch in c):
            return False
    return True


def _split_md_row(line: str) -> list[str]:
    return [c.strip() for c in line.strip().strip("|").split("|")]


def _add_md_table_to_doc(doc: Document, table_lines: list[str]):
    header = _split_md_row(table_lines[0])
    rows = [_split_md_row(x) for x in table_lines[2:]]

    ncols = len(header)
    table = doc.add_table(rows=1, cols=ncols)
    table.style = "Table Grid"

    for j in range(ncols):
        table.rows[0].cells[j].text = header[j]

    for row in rows:
        r = table.add_row().cells
        for j in range(ncols):
            r[j].text = row[j] if j < len(row) else ""


def _add_markdown_to_doc(doc: Document, text: str):
    lines = text.splitlines()
    i = 0
    buffer = []

    def flush():
        nonlocal buffer
        joined = "\n".join(buffer).strip()
        if joined:
            doc.add_paragraph(joined)
        buffer = []

    while i < len(lines):
        line = lines[i]
        if _is_md_table_line(line) and i + 1 < len(lines) and _is_md_separator_line(lines[i + 1]):
            flush()
            block = [line, lines[i + 1]]
            i += 2
            while i < len(lines) and _is_md_table_line(lines[i]):
                block.append(lines[i])
                i += 1
            _add_md_table_to_doc(doc, block)
            doc.add_paragraph("")
        else:
            buffer.append(line)
            i += 1

    flush()


def export_docx(messages):
    doc = Document()
    doc.add_heading("Rahhal CREW Output", level=1)
    doc.add_paragraph(f"Generated: {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}")

    for m in messages:
        if m.get("role") == "system":
            continue
        doc.add_heading(m["role"].capitalize(), level=2)
        _add_markdown_to_doc(doc, m["content"])

    path = "rahhal_output.docx"
    doc.save(path)
    return path


def build_final_package_override() -> str:
    return """
For this response only:

Output ONLY markdown tables. No prose outside tables.

Order:
1 System Mapping Summary
2 Objective Matrix
3 Scenario Brief as Field | Content
4 Inject Architecture
5 Facilitator Evaluation Matrix
6 Coverage Validation Summary

Use exact headers:

System Mapping Summary:
Exercise Level | Format | Hazard | Systems | Stakeholders | Escalation Authority | Duration | Constraints

Objective Matrix:
Objective | Structural Dimension | Measurable Indicator | Linked Capability

Scenario Brief:
Field | Content

Inject Architecture:
Wave | Time | Inject Format | What Participants Receive | Linked Objective | Decision Focus | Discussion Questions | What to Observe

Facilitator Evaluation Matrix:
Objective | Structural Dimension | Expected Actions | Performance Indicators | Evidence Source | Observer Notes

Coverage Validation Summary:
Item | Status | Notes

If unknown write TBC.
Max 3 discussion questions separated by semicolons.
Do not use the pipe character inside cell content.
"""


def ensure_session():
    if "msgs" not in st.session_state:
        st.session_state.msgs = [
            {"role": "system", "content": RAHHAL_SYSTEM},
            {
                "role": "assistant",
                "content": (
                    "Welcome. I am Rahhal, your Crisis Readiness Exercise Navigator.\n\n"
                    "I will guide you step by step to design a structured preparedness exercise "
                    "with governance clarity, operational alignment, and measurable evaluation outputs.\n\n"
                    "To begin, are you designing a Discussion-Based Tabletop or a Functional Exercise?"
                ),
            },
        ]
    if "pending_user" not in st.session_state:
        st.session_state.pending_user = None


def run_model_turn(user_text: str):
    clean = (user_text or "").strip()
    if not clean:
        return

    st.session_state.msgs.append({"role": "user", "content": clean})

    if clean.upper() == "FINAL PACKAGE":
        st.session_state.msgs.append({"role": "system", "content": build_final_package_override()})

    client = get_client_or_sidebar_error()
    if client is None:
        return

    try:
        response = client.chat.completions.create(
            model=DEFAULT_MODEL,
            temperature=st.session_state.get("temperature", 0.1),
            messages=st.session_state.msgs,
        )
        reply = response.choices[0].message.content
    except Exception as e:
        st.sidebar.error(f"API error: {e}")
        return

    st.session_state.msgs.append({"role": "assistant", "content": reply})


st.set_page_config(page_title="Rahhal CREW", page_icon="🧭", layout="centered")

ensure_session()

st.markdown("## Rahhal CREW")
st.markdown("Structured Crisis Readiness Exercise Navigator with clean export to Word.")
st.divider()

with st.sidebar:
    st.header("Control Panel")
    st.session_state["temperature"] = st.slider("Creativity", 0.0, 1.0, 0.1, 0.05)

    st.divider()
    if st.button("Reset chat"):
        del st.session_state["msgs"]
        st.session_state.pending_user = None
        st.rerun()

    st.divider()
    st.subheader("Package")
    if st.button("Generate full package"):
        st.session_state.pending_user = "FINAL PACKAGE"
        st.rerun()

    st.divider()
    st.subheader("Export")
    if st.button("Prepare DOCX download"):
        msgs = [m for m in st.session_state.get("msgs", []) if m.get("role") != "system"]
        if not msgs:
            st.warning("No content yet. Chat first, then export.")
        else:
            file_path = export_docx(msgs)
            with open(file_path, "rb") as f:
                st.download_button(
                    "Download Word file",
                    f,
                    file_name=file_path,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

tab_chat, tab_help = st.tabs(["Chat", "How it works"])

with tab_help:
    st.markdown(
        """
• Open the Chat tab  
• Answer the first question about exercise format  
• Provide inputs one by one as Rahhal asks  
• Click Generate full package when ready  
• Use Export in the left panel to download a Word file with real tables  
"""
    )

with tab_chat:
    for m in st.session_state.msgs:
        if m["role"] in ["assistant", "user"]:
            with st.chat_message(m["role"]):
                st.write(m["content"])

    user_input = st.chat_input(placeholder="Write your answer here")
    if user_input:
        st.session_state.pending_user = user_input
        st.rerun()

# Process pending input safely at the end of the run
if st.session_state.pending_user:
    pending = st.session_state.pending_user
    st.session_state.pending_user = None
    run_model_turn(pending)
    st.rerun()
