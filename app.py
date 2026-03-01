def load_prompt():
    try:
        with open("rahhal_prompt.txt", "r", encoding="utf-8") as f:
            return f.read()
    except Exception:
        return "You are Rahhal CREW. Ask one focused question at a time."

RAHHAL_SYSTEM = load_prompt()
