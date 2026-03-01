import os
import streamlit as st
from datetime import datetime
from docx import Document

try:
    from openai import OpenAI
except Exception:
    OpenAI = None

RAHHAL_SYSTEM = """
You are Rahhal CREW.

CREW stands for Crisis Readiness Exercise Workflow.

Purpose
Rahhal CREW is a structured human augmentation navigation system that guides users to design preparedness exercises across hazards and operational environments.
It enforces cross system alignment, structured escalation logic, and measurable evaluation outputs.

Behavior
Ask one focused question at a time.
No MCQ.
No right or wrong framing.
Use short sections and tables when outputting the final package.
If escalation authority is undefined, flag and continue and do not block progression.

Output trigger
When the user types: FINAL PACKAGE
Generate a full exercise package including:
Context summary
Objectives
Roles and responsibilities
Escalation logic
Assumptions and constraints
Inject matrix table
Evaluation approach
"""

def require_openai():
    if OpenAI is None:
        st.error("OpenAI library is not installed. Run: pip install openai")
        st.stop()

def get_client():
    require_openai()
    key = st.session_state.get("api_key", "").strip()
    if not key:
        key = os.getenv("OPENAI_API_KEY", "").strip()
    if not key:
        st.error("Add your OpenAI API key in the sidebar, or set OPENAI_API_KEY environment variable.")
        st.stop()
    return OpenAI(api_key=key)

def export_docx(chat_messages):
    doc = Document()
    doc.add_heading("Rahhal CREW Output", level=1)
    doc.add_paragraph(f"Generated: {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}")

    doc.add_heading("Conversation Log", level=2)
    for m in chat_messages:
        role = m["role"].capitalize()
        doc.add_paragraph(f"{role}: {m['content']}")

    path = "rahhal_crew_output.docx"
    doc.save(path)
    return path

st.set_page_config(page_title="Rahhal CREW", page_icon="🧭")
st.title("Rahhal CREW")
st.caption("Chatbot wrapper with DOCX export")

with st.sidebar:
    st.header("Settings")
    st.text_input("OpenAI API key", type="password", key="api_key")
    st.caption("Tip: you can also set environment variable OPENAI_API_KEY.")
    model = st.selectbox("Model", ["gpt-4.1-mini", "gpt-4.1"], index=0)
    temperature = st.slider("Creativity", 0.0, 1.0, 0.2, 0.05)
    st.divider()
    if st.button("Reset conversation"):
        st.session_state.msgs = [{"role": "system", "content": RAHHAL_SYSTEM}]
        st.rerun()

if "msgs" not in st.session_state:
    st.session_state.msgs = [{"role": "system", "content": RAHHAL_SYSTEM}]

for m in st.session_state.msgs:
    if m["role"] in ["user", "assistant"]:
        with st.chat_message(m["role"]):
            st.write(m["content"])

user_text = st.chat_input("Message Rahhal")
if user_text:
    st.session_state.msgs.append({"role": "user", "content": user_text})
    with st.chat_message("user"):
        st.write(user_text)

    client = get_client()
    with st.chat_message("assistant"):
        try:
            resp = client.chat.completions.create(
                model=model,
                temperature=temperature,
                messages=st.session_state.msgs
            )
            answer = resp.choices[0].message.content
        except Exception as e:
            st.error(f"API error: {e}")
            st.stop()

        st.write(answer)

    st.session_state.msgs.append({"role": "assistant", "content": answer})

st.divider()
col1, col2 = st.columns(2)

with col1:
    if st.button("Export DOCX"):
        msgs = [m for m in st.session_state.msgs if m["role"] != "system"]
        path = export_docx(msgs)
        with open(path, "rb") as f:
            st.download_button(
                "Download DOCX",
                data=f,
                file_name=path,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

with col2:
    st.caption("Type FINAL PACKAGE in chat when you want the full structured output.")
