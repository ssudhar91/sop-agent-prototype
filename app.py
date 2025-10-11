# app.py
import streamlit as st
import pandas as pd
import ast
from typing import List

st.set_page_config(page_title="Novotech SOP Finder", layout="wide")

st.title("Novotech — SOP Finder (Prototype)")

# --- Upload or use repo file ---
uploaded = st.file_uploader("Upload SOP Excel (or leave empty to use repo file)", type=["xlsx","xls","csv"])
if uploaded is not None:
    df_raw = pd.read_excel(uploaded, sheet_name=0)
else:
    # fallback path if deployed with data in repo
    try:
        df_raw = pd.read_excel("data/Novotech_SOP_Matrix.xlsx", sheet_name=0)
    except Exception as e:
        st.warning("No file found. Please upload your SOP Excel or add it to data/Novotech_SOP_Matrix.xlsx in repo.")
        st.stop()

st.sidebar.header("Parsing options")
sep = st.sidebar.text_input("Multi-value separator (if roles/timelines are '2;3')", value=";")
role_col = st.sidebar.text_input("Role header column name (exact)", value="Role")
notes_col = st.sidebar.text_input("Notes column name (exact)", value="Notes")
timeline_cols = st.sidebar.text_input("Timeline columns (comma-separated names for columns that have 1,2,3 values)", value="Onboarding")
# Adjust these defaults to match your sheet

# --- Helper parsing functions ---
def split_multi(val, sep=";"):
    if pd.isna(val):
        return []
    # if already a list-like string e.g. "['2','3']" attempt literal_eval
    if isinstance(val, str) and (val.strip().startswith('[') or ',' in val and sep not in val):
        try:
            return list(ast.literal_eval(val))
        except Exception:
            pass
    # split on sep
    if isinstance(val, str):
        parts = [x.strip() for x in val.split(sep) if x.strip()!=""]
        return parts
    if isinstance(val, (int, float)):
        return [str(int(val))]
    if isinstance(val, list):
        return [str(x) for x in val]
    return []

# --- Preprocess DataFrame ---
df = df_raw.copy()

# Normalize column names
df.columns = [str(c).strip() for c in df.columns]

# Example: find common columns
timeline_column = None
for c in df.columns:
    if any(x.lower() in c.lower() for x in ["timeline","onboarding","completion","when"]):
        timeline_column = c
        break
if timeline_column is None and timeline_cols:
    timeline_column = timeline_cols

# Create normalized timeline list column
if timeline_column in df.columns:
    df["_timeline_list"] = df[timeline_column].apply(lambda v: split_multi(v, sep=sep))
else:
    df["_timeline_list"] = [[] for _ in range(len(df))]

# If there's a roles area across many columns (roles as headers), detect numeric codes 1/2/3 in cells
# Suppose your sheet uses rows as SOPs and columns from col E onward are role headers with '1','2','3' marks
# We make a long table where each SOP x Role with a mark becomes a row.
role_headers = []
for c in df.columns:
    # heuristic: role headers are likely not the SOP title or notes; look for column names that are not 'Title'/'Notes'
    if c.lower() not in ["title","sop","sop title","notes","notes/region","description"]:
        # also skip timeline column
        if c != timeline_column and c != role_col:
            role_headers.append(c)

# Build long-format mapping if role_headers have numeric marks
long_rows = []
for idx, row in df.iterrows():
    sop_title = row.get("Title") or row.get("SOP") or row.get(df.columns[0])
    for rh in role_headers:
        val = row.get(rh)
        if pd.isna(val): 
            continue
        sval = str(val).strip()
        if sval=="":
            continue
        # if the cell contains numbers like "1" or "2;3"
        marks = split_multi(sval, sep=sep)
        for m in marks:
            long_rows.append({
                "SOP Title": sop_title,
                "Role": rh,
                "Mark": m,
                "Notes": row.get(notes_col, ""),
                "TimelineList": row.get("_timeline_list", []),
                "RowIndex": idx
            })

long_df = pd.DataFrame(long_rows)

st.sidebar.header("Filters")
all_roles = sorted(long_df["Role"].unique().tolist())
selected_roles = st.sidebar.multiselect("Select roles", options=all_roles, default=all_roles[:5])
selected_marks = st.sidebar.multiselect("Completion codes to include (1,2,3)", options=["1","2","3"], default=["1","2","3"])
region_filter = st.sidebar.text_input("Region (leave empty for all, e.g., China, Korea, Taiwan)", value="")

# Apply filters
filtered = long_df[long_df["Role"].isin(selected_roles)]
filtered = filtered[filtered["Mark"].isin(selected_marks)]
if region_filter.strip():
    region_lower = region_filter.strip().lower()
    filtered = filtered[filtered[filtered["Notes"].str.lower().fillna("").str.contains(region_lower)]

# Show table
st.subheader("Filtered SOPs")
st.write(f"Rows: {len(filtered)}")
st.dataframe(filtered.reset_index(drop=True))

# Optional: natural language question about the filtered set (requires LLM API)
st.subheader("Ask a question about the filtered SOPs (optional)")
user_q = st.text_input("Type your question here (e.g., 'Which SOPs must be done within 2 weeks for Role X?')")

use_llm = st.checkbox("Use LLM to generate natural language answer (requires API key in secrets)", value=False)
if user_q and use_llm:
    # Prepare context: small table snippet
    snippet = filtered.head(50).to_dict(orient="records")
    # Call LLM - provider-specific code (placeholder)
    st.info("LLM integration not configured. See README to add your OPENAI_API_KEY to secrets and uncomment LLM block.")
    # Example: If you configure openai, you would send the question + snippet as context to the model.
