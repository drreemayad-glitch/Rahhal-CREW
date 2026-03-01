import os
import streamlit as st
from datetime import datetime
from docx import Document

try:
    from openai import OpenAI
except Exception:
    OpenAI = None


# =========================
# CONFIG
# =========================

st.set_page_config(page_title="Rahhal CREW", page_icon="⚙️", layout="centered")

DEFAULT_MODEL = "gpt-4.1-mini"


# =========================
# LOAD SYSTEM PROMPT
# =========================

def load_prompt():
    try:
        with open("rahhal_prompt.txt", "r", encoding="utf-8") as f:
            return f.read()
    except Exception:
        return "Rahhal CREW system active."


RAHHAL_SYSTEM = load_prompt()


# =========================
# OPENAI CLIENT
# =========================

def get_client():
    if OpenAI is None:
        st.sidebar.error("OpenAI library missing.")
        return None

    key = os.getenv("OPENAI_API_KEY")
    if not key:
        st.sidebar.error("OPENAI_API_KEY not configured in Streamlit Secrets.")
        return None

    return OpenAI(api_key=key)


# =========================
# DOCX EXPORT
# =========================

def export_docx(messages):
    doc = Document()
    doc.add_heading("Rahhal CREW Output", level=1)
    doc.add_paragraph(f"Generated: {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}")

    for m in messages:
        if m["role"] == "system":
            continue
        doc.add_heading(m["role"].capitalize(), level=2)
        doc.add_paragraph(m["content"])

    path = "rahhal_output.docx"
    doc.save(path)
    return path


# =========================
# SESSION INIT
# =========================

if "messages" not in st.session_state:
    st.session_state.messages = [
        {"role": "system", "content": RAHHAL_SYSTEM},
        {
            "role": "assistant",
            "content": (
                "Rahhal CREW activated.\n\n"
                "Structured crisis readiness exercise design system.\n\n"
                "Select exercise format:\n"
                "Discussion-Based Tabletop or Functional Exercise?"
            ),
        },
    ]

if "pending_input" not in st.session_state:
    st.session_state.pending_input = None


# =========================
# HEADER
# =========================

st.markdown("## Rahhal CREW")
st.markdown("Structured Crisis Readiness Exercise Design System")
st.divider()


# =========================
# SIDEBAR
# =========================

with st.sidebar:

    st.header("Control Panel")

    temperature = st.slider("Creativity", 0.0, 1.0, 0.1, 0.05)

    st.divider()

    if st.button("Reset session"):
        del st.session_state.messages
        st.session_state.pending_input = None
        st.rerun()

    st.divider()

    if st.button("Generate full package"):
        st.session_state.pending_input = "FINAL PACKAGE"
        st.rerun()

    st.divider()

    if st.button("Prepare DOCX export"):
        msgs = [m for m in st.session_state.messages if m["role"] != "system"]
        if msgs:
            file_path = export_docx(msgs)
            with open(file_path, "rb") as f:
                st.download_button(
                    "Download Word file",
                    f,
                    file_name=file_path,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
        else:
            st.warning("No content to export.")


# =========================
# TABS
# =========================

tab_chat, tab_help = st.tabs(["Chat", "How it works"])


with tab_help:
    st.markdown(
        """
• Open the Chat tab  
• Select exercise format  
• Respond step by step  
• Click Generate full package when ready  
• Export Word document from sidebar  
"""
    )


# =========================
# CHAT TAB
# =========================

with tab_chat:

    # Render messages with neutral avatars
    for m in st.session_state.messages:
        if m["role"] == "assistant":
            with st.chat_message("assistant", avatar="⚙️"):
                st.write(m["content"])
        elif m["role"] == "user":
            with st.chat_message("user", avatar="👤"):
                st.write(m["content"])

    user_text = st.chat_input("Write your response")

    if user_text:
        st.session_state.pending_input = user_text
        st.rerun()


# =========================
# PROCESS MODEL TURN SAFELY
# =========================

if st.session_state.pending_input:

    user_text = st.session_state.pending_input
    st.session_state.pending_input = None

    st.session_state.messages.append(
        {"role": "user", "content": user_text}
    )

    client = get_client()
    if client:

        try:
            response = client.chat.completions.create(
                model=DEFAULT_MODEL,
                temperature=temperature,
                messages=st.session_state.messages,
            )

            reply = response.choices[0].message.content

            st.session_state.messages.append(
                {"role": "assistant", "content": reply}
            )

        except Exception as e:
            st.sidebar.error(f"API error: {e}")

    st.rerun()
