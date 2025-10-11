# app.py
import streamlit as st
import pandas as pd
import os
import openai

# -------------------------
# Config
# -------------------------
DATA_PATH = "data/Novotech_SOP_Matrix.xlsx"  # fallback Excel
OUTPUT_DIR = "output"

# -------------------------
# Streamlit UI
# -------------------------
st.set_page_config(page_title="Novotech SOP Finder", layout="wide")
st.title("Novotech SOP Finder with GPT-5")

uploaded = st.file_uploader("Upload SOP Excel (optional, else repo file used)", type=["xlsx"])
if uploaded is not None:
    df = pd.read_excel(uploaded, sheet_name=0, header=None)
else:
    try:
        df = pd.read_excel(DATA_PATH, sheet_name=0, header=None)
    except Exception as e:
        st.error("No Excel file found. Upload or place Novotech_SOP_Matrix.xlsx in data/ folder.")
        st.stop()

# -------------------------
# Parse Excel according to your structure
# -------------------------
# Role names (row 3, col E onwards)
roles = df.iloc[2, 4:].tolist()

# Group names (row 1, col E onwards)
groups = df.iloc[0, 4:].tolist()

# SOP rows start from row 4
sop_rows = df.iloc[3:, :]

# Build a long-format DataFrame: one row per SOP per role if apply_flag in 1/2/3
records = []
for col_idx, role in enumerate(roles, start=4):
    group_name = groups[col_idx - 4]
    for row_idx, row in sop_rows.iterrows():
        apply_flag = row[col_idx]
        if apply_flag in [1, 2, 3]:
            records.append({
                "Role": role,
                "Group": group_name,
                "Business Unit": row[0],
                "SOP Type": row[1],
                "SOP Code": row[2],
                "Title": row[3],
                "Apply Level": apply_flag
            })

long_df = pd.DataFrame(records)

# -------------------------
# Sidebar Filters
# -------------------------
st.sidebar.header("Filters")
all_roles = sorted(long_df["Role"].unique())
selected_roles = st.sidebar.multiselect("Select roles", all_roles, default=all_roles[:5])
selected_marks = st.sidebar.multiselect("Apply Level (1,2,3)", options=[1,2,3], default=[1,2,3])
selected_groups = st.sidebar.multiselect("Select groups", long_df["Group"].unique(), default=long_df["Group"].unique())

filtered_df = long_df[
    (long_df["Role"].isin(selected_roles)) &
    (long_df["Apply Level"].isin(selected_marks)) &
    (long_df["Group"].isin(selected_groups))
]

st.subheader("Filtered SOPs")
st.write(f"Total rows: {len(filtered_df)}")
st.dataframe(filtered_df.reset_index(drop=True))

# -------------------------
# LLM Question Answering
# -------------------------
st.subheader("Ask a question about the filtered SOPs")

user_question = st.text_input("Type your question here:")

if user_question:
    if "OPENAI_API_KEY" not in st.secrets:
        st.error("Add your OPENAI_API_KEY in Streamlit Secrets to enable GPT-5.")
    else:
        openai.api_key = st.secrets["OPENAI_API_KEY"]
        # Prepare context: first 50 rows
        context = filtered_df.head(50).to_dict(orient="records")
        prompt = f"Answer the question based on the SOP context below:\n\nContext:\n{context}\n\nQuestion: {user_question}\nAnswer clearly."
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
