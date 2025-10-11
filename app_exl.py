import os
import streamlit as st
import pandas as pd

# -------------------------------
# Streamlit Setup
# -------------------------------
st.set_page_config(page_title="Agentic AI – Excel SOP Query (No LLM)", layout="wide")
st.title("Agentic AI – Excel SOP Query (No LLM)")

# -------------------------------
# Path to master Excel file
# -------------------------------
base_dir = os.path.dirname(__file__)
excel_file_path = os.path.join(base_dir, "data", "Novotech_SOP_Matrix.xlsx")

if not os.path.exists(excel_file_path):
    st.error(f"Master Excel file not found at {excel_file_path}")
else:
    st.write(f"Using master Excel file: {excel_file_path}")

    # -------------------------------
    # Load Excel file and flatten sheets into text per role
    # -------------------------------
    xls = pd.ExcelFile(excel_file_path)
    role_texts = {}
    for sheet in xls.sheet_names:
        df = xls.parse(sheet)
        role_texts[sheet] = "\n".join(df.astype(str).values.flatten())

    # -------------------------------
    # User selects role
    # -------------------------------
    roles = list(role_texts.keys())
    selected_role = st.selectbox("Select a role to query:", roles)

    query = st.text_input("Ask a question about SOP:")

    if query:
        text = role_texts[selected_role]
        lower_query = query.lower()

        # -------------------------------
        # Simple keyword-based search
        # -------------------------------
        matching_lines = [line for line in text.split("\n") if lower_query in line.lower()]

        st.markdown("**Agentic AI says (Keyword Match Mode):**")
        if matching_lines:
            # Show top 5 matching lines
            for line in matching_lines[:5]:
                st.write(f"- {line}")
        else:
            st.write("No matching information found in this role's SOP.")
