import os
import streamlit as st
from datetime import datetime
from docx import Document

st.set_page_config(page_title="Rahhal CREW", page_icon="🧭")
st.write("App started OK")

try:
    from openai import OpenAI
except Exception:
    OpenAI = None


def load_prompt():
    try:
        with open("rahhal_prompt.txt", "r", encoding="utf-8") as f:
            return f.read()
    except Exception:
        return "You are Rahhal CREW. Ask one focused question at a time."


RAHHAL_SYSTEM = load_prompt()


def get_client():
    if OpenAI is None:
        st.error("OpenAI library not installed.")
        st.stop()

    key = st.session_state.get("api_key", "").strip()
    if not key:
        key = os.getenv("OPENAI_API_KEY", "").strip()

    if not key:
        st.warning("Add your OpenAI API key in the sidebar.")
        st.stop()

    return OpenAI(api_key=key)


# ---------------- DOCX TABLE PARSER ---------------- #

def _is_md_table_line(line):
    s = line.strip()
    return s.startswith("|") and s.endswith("|") and s.count("|") >= 3


def _is_md_separator_line(line):
    s = line.strip().replace(" ", "")
    if not (s.startswith("|") and s.endswith("|")):
        return False
    cells = [c for c in s.strip("|").split("|")]
    for c in cells:
        if c and not all(ch in "-:" for ch in c):
            return False
    return True


def _split_md_row(line):
    return [c.strip() for c in line.strip().strip("|").split("|")]


def _add_md_table_to_doc(doc, table_lines):
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
            if j < len(row):
                r[j].text = row[j]


def _add_markdown_to_doc(doc, text):
    lines = text.splitlines()
    i = 0
    buffer = []

    def flush():
        nonlocal buffer
        if buffer:
            doc.add_paragraph("\n".join(buffer))
            buffer = []

    while i < len(lines):
        line = lines[i]

        if (
            _is_md_table_line(line)
            and i + 1 < len(lines)
            and _is_md_separator_line(lines[i + 1])
        ):
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
    doc.add_paragraph(
        f"Generated: {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}"
    )

    for m in messages:
        doc.add_heading(m["role"].capitalize(), level=2)
        _add_markdown_to_doc(doc, m["content"])

    path = "rahhal_output.docx"
    doc.save(path)
    return path


# ---------------- UI ---------------- #

st.title("Rahhal CREW")
st.caption("Crisis Readiness Exercise Workflow Engine")

with st.sidebar:
    st.header("Settings")
    st.text_input("OpenAI API key", type="password", key="api_key")
    model = st.selectbox("Model", ["gpt-4.1-mini", "gpt-4.1"])
    temperature = st.slider("Creativity", 0.0, 1.0, 0.1)

    if st.button("Reset"):
        if "msgs" in st.session_state:
            del st.session_state["msgs"]
        st.rerun()


if "msgs" not in st.session_state:
    st.session_state.msgs = [
        {"role": "system", "content": RAHHAL_SYSTEM},
        {
            "role": "assistant",
            "content": "Are you designing a Discussion-Based Tabletop or a Functional Exercise?",
        },
    ]


for m in st.session_state.msgs:
    if m["role"] in ["assistant", "user"]:
        with st.chat_message(m["role"]):
            st.write(m["content"])


user_input = st.chat_input("Message Rahhal")

if user_input:
    st.session_state.msgs.append({"role": "user", "content": user_input})

    with st.chat_message("user"):
        st.write(user_input)

    client = get_client()

    with st.chat_message("assistant"):
        try:
            response = client.chat.completions.create(
                model=model,
                temperature=temperature,
                messages=st.session_state.msgs,
            )
            reply = response.choices[0].message.content
        except Exception as e:
            st.error(f"API error: {e}")
            st.stop()

        st.write(reply)

    st.session_state.msgs.append({"role": "assistant", "content": reply})


st.divider()

if st.button("Export DOCX"):
    msgs = [m for m in st.session_state.msgs if m["role"] != "system"]
    file_path = export_docx(msgs)

    with open(file_path, "rb") as f:
        st.download_button(
            "Download DOCX",
            f,
            file_name=file_path,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

