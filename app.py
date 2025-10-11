# app.py
import streamlit as st
import os
import openai

# -------------------------
# Config
# -------------------------
OUTPUT_DIR = "output"

st.set_page_config(page_title="Novotech SOP Finder - Agentic AI", layout="wide")
st.title("Novotech SOP Finder — Agentic AI with GPT-5")

# -------------------------
# Load preprocessed SOP files
# -------------------------
all_files = [f for f in os.listdir(OUTPUT_DIR) if f.endswith(".txt")]
if not all_files:
    st.error("No preprocessed SOP files found in 'output/' folder. Upload them first.")
    st.stop()

# Sidebar: select roles
st.sidebar.header("Select Roles")
selected_roles = st.sidebar.multiselect("Choose roles to include", options=all_files, default=all_files[:5])

# Read selected SOP files
sop_context = []
for role_file in selected_roles:
    file_path = os.path.join(OUTPUT_DIR, role_file)
    with open(file_path, "r", encoding="utf-8") as f:
        lines = f.readlines()
        # skip empty lines
        lines = [l.strip() for l in lines if l.strip()]
        sop_context.extend(lines)

if not sop_context:
    st.warning("No SOP lines found for selected roles.")
    st.stop()

# Display SOP lines (first 100)
st.subheader("SOP Lines for Selected Roles")
st.text("\n".join(sop_context[:100]))

# -------------------------
# GPT-5 Question Answering
# -------------------------
st.subheader("Ask a question about the SOPs")
user_question = st.text_input("Type your question here:")

if user_question:
    if "OPENAI_API_KEY" not in st.secrets:
        st.error("Please add your OPENAI_API_KEY in Streamlit secrets to enable GPT-5.")
    else:
        openai.api_key = st.secrets["OPENAI_API_KEY"]
        # Prepare prompt
        prompt = f"Answer the question based on the SOP context below:\n\nContext:\n" \
                 f"{sop_context[:200]}\n\nQuestion: {user_question}\nAnswer clearly."
        try:
            response = openai.ChatCompletion.create(
                model="gpt-5-mini",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.2
            )
            answer = response['choices'][0]['message']['content']
            st.markdown(f"**GPT-5 Answer:** {answer}")
        except Exception as e:
            st.error(f"Error calling GPT-5 API: {e}")
