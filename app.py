import os
import streamlit as st
from datetime import datetime
from docx import Document

from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.section import WD_ORIENT

try:
    from openai import OpenAI
except Exception:
    OpenAI = None


DEFAULT_MODEL = "gpt-4.1-mini"


def load_prompt() -> str:
    try:
        with open("rahhal_prompt.txt", "r", encoding="utf-8") as f:
            return f.read()
    except Exception:
        return "Rahhal CREW system active. Ask one focused question at a time."


RAHHAL_SYSTEM = load_prompt()


def get_client_or_sidebar_error():
    if OpenAI is None:
        st.sidebar.error("OpenAI library missing. Check requirements.txt.")
        return None

    key = (os.getenv("OPENAI_API_KEY") or "").strip()
    if not key:
        st.sidebar.error("Missing OPENAI_API_KEY. Add it in Streamlit Secrets.")
        return None

    return OpenAI(api_key=key)


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
Keep each cell concise. Avoid paragraphs inside a cell.
"""


def ensure_session():
    if "messages" not in st.session_state:
        st.session_state.messages = [
            {"role": "system", "content": RAHHAL_SYSTEM},
            {
                "role": "assistant",
                "content": (
                    "Rahhal CREW activated.\n\n"
                    "Structured crisis readiness exercise design system.\n\n"
                    "Select exercise format:\n"
                    "Discussion Based Tabletop or Functional Exercise?"
                ),
            },
        ]
    if "pending_input" not in st.session_state:
        st.session_state.pending_input = None
    if "temperature" not in st.session_state:
        st.session_state.temperature = 0.1


def _is_md_table_line(line: str) -> bool:
    s = line.strip()
    return s.startswith("|") and s.endswith("|") and s.count("|") >= 3


def _is_md_separator_line(line: str) -> bool:
    s = line.strip().replace(" ", "")
    if not (s.startswith("|") and s.endswith("|")):
        return False
    cells = [c for c in s.strip("|").split("|")]
    if not cells:
        return False
    for c in cells:
        if c == "":
            continue
        ok = all(ch in "-:" for ch in c)
        if not ok:
            return False
        if "-" not in c:
            return False
    return True


def _split_md_row(line: str) -> list[str]:
    return [c.strip() for c in line.strip().strip("|").split("|")]


def _style_cell(cell, bold=False, center=False, font_pt=8):
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    for p in cell.paragraphs:
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
        for run in p.runs:
            run.font.size = Pt(font_pt)
            run.bold = bold


def _add_md_table_to_doc(doc: Document, table_lines: list[str], usable_width_in=9.0, font_pt=8):
    if len(table_lines) < 2:
        return

    header = _split_md_row(table_lines[0])
    rows = [_split_md_row(x) for x in table_lines[2:]]

    ncols = max(1, len(header))
    header = (header + [""] * (ncols - len(header)))[:ncols]

    fixed_rows = []
    for r in rows:
        rr = (r + [""] * (ncols - len(r)))[:ncols]
        fixed_rows.append(rr)

    table = doc.add_table(rows=1, cols=ncols)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False

    usable_width = Inches(usable_width_in)
    col_width = usable_width / ncols

    for j in range(ncols):
        table.columns[j].width = col_width

    for j in range(ncols):
        c = table.rows[0].cells[j]
        c.width = col_width
        c.text = header[j]
        _style_cell(c, bold=True, center=True, font_pt=font_pt)

    for r in fixed_rows:
        row_cells = table.add_row().cells
        for j in range(ncols):
            row_cells[j].width = col_width
            row_cells[j].text = r[j]
            _style_cell(row_cells[j], bold=False, center=False, font_pt=font_pt)


def _add_markdown_to_doc(doc: Document, text: str):
    lines = text.splitlines()
    i = 0
    para_buf = []

    def flush_para():
        nonlocal para_buf
        joined = "\n".join(para_buf).strip()
        if joined:
            p = doc.add_paragraph(joined)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(6)
        para_buf = []

    while i < len(lines):
        line = lines[i]

        if _is_md_table_line(line) and i + 1 < len(lines) and _is_md_separator_line(lines[i + 1]):
            flush_para()
            block = [line, lines[i + 1]]
            i += 2
            while i < len(lines) and _is_md_table_line(lines[i]):
                block.append(lines[i])
                i += 1
            _add_md_table_to_doc(doc, block, usable_width_in=9.0, font_pt=8)
            doc.add_paragraph("")
            continue

        para_buf.append(line)
        i += 1

    flush_para()


def export_docx(messages):
    doc = Document()

    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)
    section.top_margin = Inches(0.6)
    section.bottom_margin = Inches(0.6)

    doc.add_paragraph(f"Generated: {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}")
    doc.add_paragraph("")

    for m in messages:
        if m.get("role") == "system":
            continue

        role = m.get("role", "").capitalize()
        if role:
            h = doc.add_paragraph(role)
            for run in h.runs:
                run.bold = True

        _add_markdown_to_doc(doc, m.get("content", ""))
        doc.add_paragraph("")

    path = "rahhal_output.docx"
    doc.save(path)
    return path


def run_model_turn(user_text: str):
    clean = (user_text or "").strip()
    if not clean:
        return

    st.session_state.messages.append({"role": "user", "content": clean})

    if clean.upper() == "FINAL PACKAGE":
        st.session_state.messages.append({"role": "system", "content": build_final_package_override()})

    client = get_client_or_sidebar_error()
    if client is None:
        return

    try:
        resp = client.chat.completions.create(
            model=DEFAULT_MODEL,
            temperature=float(st.session_state.temperature),
            messages=st.session_state.messages,
        )
        reply = resp.choices[0].message.content
        st.session_state.messages.append({"role": "assistant", "content": reply})
    except Exception as e:
        st.sidebar.error(f"API error: {e}")


st.set_page_config(page_title="Rahhal CREW", page_icon="🛡", layout="wide")
ensure_session()

st.markdown("# Rahhal CREW")
st.markdown("AI Augmented Crisis Readiness Exercise Design System")
st.divider()

with st.sidebar:
    st.header("Control Panel")
    st.session_state.temperature = st.slider(
        "Creativity", 0.0, 1.0, float(st.session_state.temperature), 0.05
    )

    st.divider()

    if st.button("Reset session"):
        if "messages" in st.session_state:
            del st.session_state["messages"]
        st.session_state.pending_input = None
        st.rerun()

    st.divider()

    if st.button("Generate structured exercise package"):
        st.session_state.pending_input = "FINAL PACKAGE"
        st.rerun()

    st.divider()

    if st.button("Prepare Word export"):
        msgs = [m for m in st.session_state.messages if m.get("role") != "system"]
        if not msgs:
            st.warning("No content to export yet.")
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
• Select exercise format  
• Respond step by step as Rahhal asks  
• Click Generate structured exercise package when ready  
• Export the Word package from the left panel  
"""
    )

with tab_chat:
    for m in st.session_state.messages:
        if m["role"] == "assistant":
            with st.chat_message("assistant", avatar="🛡"):
                st.write(m["content"])
        elif m["role"] == "user":
            with st.chat_message("user", avatar="👤"):
                st.write(m["content"])

    user_text = st.chat_input(placeholder="Write your answer here")
    if user_text:
        st.session_state.pending_input = user_text
        st.rerun()

if st.session_state.pending_input:
    pending = st.session_state.pending_input
    st.session_state.pending_input = None
    run_model_turn(pending)
    st.rerun()
