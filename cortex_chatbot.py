import streamlit as st
from snowflake.snowpark.context import get_active_session
import json

# Page configuration
st.set_page_config(
    page_title="Snowflake Cortex Chatbot",
    page_icon="❄️",
    layout="wide",
)

# Sidebar configuration
with st.sidebar:
    st.title("❄️ Cortex Chatbot")
    st.divider()

    model = st.selectbox(
        "Model",
        options=[
            "snowflake-arctic",
            "mistral-large2",
            "llama3.1-70b",
            "llama3.1-8b",
            "llama3.2-1b",
            "llama3.2-3b",
            "mistral-7b",
            "gemma-7b",
            "mixtral-8x7b",
            "reka-flash",
            "reka-core",
        ],
        index=1,
    )

    temperature = st.slider("Temperature", min_value=0.0, max_value=1.0, value=0.7, step=0.05)
    max_tokens = st.slider("Max tokens", min_value=100, max_value=4096, value=1024, step=100)

    st.divider()
    system_prompt = st.text_area(
        "System prompt",
        value="You are a helpful assistant. Answer questions clearly and concisely.",
        height=120,
    )

    st.divider()
    if st.button("Clear conversation", use_container_width=True):
        st.session_state.messages = []
        st.rerun()

    st.caption(f"Messages in history: {len(st.session_state.get('messages', []))}")


def get_snowflake_session():
    try:
        return get_active_session()
    except Exception:
        st.error(
            "Could not connect to Snowflake. Make sure this app is running inside Streamlit in Snowflake (SiS)."
        )
        st.stop()


def call_cortex_complete(session, model: str, messages: list, system_prompt: str, temperature: float, max_tokens: int) -> str:
    """Call Snowflake Cortex COMPLETE function with conversation history."""
    conversation = [{"role": "system", "content": system_prompt}] + messages

    # Escape single quotes for SQL embedding
    messages_json = json.dumps(conversation).replace("'", "''")

    options_json = json.dumps({"temperature": temperature, "max_tokens": max_tokens}).replace("'", "''")

    sql = f"""
        SELECT SNOWFLAKE.CORTEX.COMPLETE(
            '{model}',
            PARSE_JSON('{messages_json}'),
            PARSE_JSON('{options_json}')
        ) AS response
    """

    result = session.sql(sql).collect()
    raw = result[0]["RESPONSE"]

    # COMPLETE returns a JSON string; extract the assistant message content
    if isinstance(raw, str):
        try:
            parsed = json.loads(raw)
            return parsed["choices"][0]["messages"]
        except (json.JSONDecodeError, KeyError, IndexError):
            return raw
    return str(raw)


# ── Main chat UI ──────────────────────────────────────────────────────────────

st.title("❄️ Snowflake Cortex Chatbot")

# Initialise message history
if "messages" not in st.session_state:
    st.session_state.messages = []

# Render existing messages
for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

# Chat input
if prompt := st.chat_input("Ask me anything…"):
    # Show the user's message immediately
    with st.chat_message("user"):
        st.markdown(prompt)
    st.session_state.messages.append({"role": "user", "content": prompt})

    # Call Cortex and stream the response
    with st.chat_message("assistant"):
        with st.spinner("Thinking…"):
            session = get_snowflake_session()
            try:
                response = call_cortex_complete(
                    session=session,
                    model=model,
                    messages=st.session_state.messages,
                    system_prompt=system_prompt,
                    temperature=temperature,
                    max_tokens=max_tokens,
                )
                st.markdown(response)
                st.session_state.messages.append({"role": "assistant", "content": response})
            except Exception as e:
                error_msg = f"Error calling Cortex: {e}"
                st.error(error_msg)
