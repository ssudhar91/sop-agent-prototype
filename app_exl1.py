import os
import io
import streamlit as st
import pandas as pd

# -------------------------------
# Streamlit Setup
# -------------------------------
st.set_page_config(page_title="Novotech SOP Matrix", layout="wide")
st.title("Novotech SOP Matrix")

# -------------------------------
# Path to master Excel file
# -------------------------------
base_dir = os.path.dirname(__file__)
excel_file_path = os.path.join(base_dir, "data", "Novotech_SOP_Matrix.xlsx")

if not os.path.exists(excel_file_path):
    st.error(f"Master Excel file not found at {excel_file_path}")
    st.stop()

# -------------------------------
# Settings
# -------------------------------
HEADER_GROUP_ROW = 0
HEADER_COLS_ROW = 2
DATA_START_ROW = 3

category_map = {
    "Within 2 weeks": 1,
    "Within 90 days": 2,
    "Before task": 3
}

regions = ["china", "korea", "taiwan", "hong kong", "india", "us", "uk"]

# -------------------------------
# Helper functions
# -------------------------------
def pick_column(df, candidates):
    cols_lower = {c.lower(): c for c in df.columns}
    for cand in candidates:
        if cand and cand.lower() in cols_lower:
            return cols_lower[cand.lower()]
    return None

def to_int_safe(v):
    try:
        if pd.isna(v):
            return None
        if isinstance(v, str):
            v = v.strip()
            if v == "":
                return None
            return int(float(v))
        return int(v)
    except Exception:
        return None

def detect_regions(text):
    text = str(text).lower()
    found = [r for r in regions if r in text]
    return ", ".join(found) if found else ""

def checkbox_group(label: str, values, key_prefix: str):
    st.markdown(f"**{label}**")
    selected = []
    for v in values:
        v_display = "(blank)" if (pd.isna(v) or str(v).strip() == "") else str(v)
        if st.checkbox(v_display, value=True, key=f"{key_prefix}_{v_display}"):
            selected.append(v)
    return selected

# -------------------------------
# Load workbook
# -------------------------------
xls = pd.ExcelFile(excel_file_path)
sheets = xls.sheet_names
sheet_choice = st.selectbox("Choose sheet:", sheets)

raw = pd.read_excel(excel_file_path, sheet_name=sheet_choice, header=None, dtype=object)
if raw.shape[0] <= HEADER_COLS_ROW:
    st.error("The selected sheet doesn't have the expected header rows.")
    st.stop()

group_row = raw.iloc[HEADER_GROUP_ROW].copy().fillna(method="ffill")
header_row = raw.iloc[HEADER_COLS_ROW].astype(str).tolist()
data_df = raw.iloc[DATA_START_ROW:].copy().reset_index(drop=True)
data_df.columns = header_row

if len(header_row) <= 4:
    st.error("Unable to detect role columns beyond column E.")
    st.stop()

# Build group-to-role mapping
groups_map = {}
for col_idx, col_name in enumerate(header_row):
    if col_idx < 4:
        continue
    raw_group_val = group_row.iloc[col_idx] if col_idx < len(group_row) else ""
    group_name = str(raw_group_val).strip() if not pd.isna(raw_group_val) else ""
    if group_name == "" or group_name.lower() in ("nan", "none"):
        group_name = "Ungrouped"
    groups_map.setdefault(group_name, []).append((col_idx, col_name))
groups_sorted = sorted(groups_map.keys())

# -------------------------------
# UI selections
# -------------------------------
selected_group = st.selectbox("Choose the Group:", groups_sorted)

role_entries = groups_map.get(selected_group, [])
display_to_colidx = {}
for col_idx, role_name in role_entries:
    display_to_colidx[f"{role_name} (col {col_idx})"] = col_idx

selected_role_display = st.selectbox("Choose the Role (within group):", list(display_to_colidx.keys()))
selected_col_idx = display_to_colidx[selected_role_display]

sop_category = st.selectbox("Choose the SOP category:", list(category_map.keys()))
category_value = category_map[sop_category]

# -------------------------------
# Identify key columns
# -------------------------------
practice_col = pick_column(data_df, ["Practice", "Department", "Function"])
group_col = pick_column(data_df, ["Group", "SOP Group", "Team Group"])
bu_col = pick_column(data_df, ["Business Unit", "BusinessUnit"])
sop_type_col = pick_column(data_df, ["SOP Type", "Type"])
number_col = pick_column(data_df, ["Number", "SOP Number", "No", "ID"])
title_col = pick_column(data_df, ["Title", "SOP Title"])
notes_col = pick_column(data_df, ["Notes", "Remarks", "Comments", "Region Notes"])

if title_col is None and len(data_df.columns) >= 4:
    title_col = data_df.columns[3]

# -------------------------------
# Apply main role/category filter
# -------------------------------
role_series = data_df.iloc[:, selected_col_idx].apply(to_int_safe)
filtered = data_df[role_series == category_value].copy()

if notes_col:
    filtered["RegionsDetected"] = filtered[notes_col].apply(detect_regions)
else:
    filtered["RegionsDetected"] = ""

# -------------------------------
# Sidebar filters: Practice, Group, Category, Region
# -------------------------------
with st.sidebar:
    st.header("Filters")

    # Practice filter
    if practice_col and practice_col in filtered.columns:
        practice_vals = sorted(filtered[practice_col].dropna().unique())
        selected_practice = checkbox_group("Practice", practice_vals, "practice")
    else:
        selected_practice = None

    # Group filter
    group_vals = sorted(groups_map.keys())
    selected_group_filter = checkbox_group("Group", group_vals, "grpfilter")

    # Category filter
    cat_vals = list(category_map.keys())
    selected_cat_filter = checkbox_group("Category", cat_vals, "catfilter")

    # Region filter
    region_present = set()
    for r in regions:
        if any(filtered["RegionsDetected"].str.contains(r, na=False)):
            region_present.add(r)
    region_present = sorted(region_present)
    selected_region = checkbox_group("Region", region_present, "regionfilter") if region_present else None

# -------------------------------
# Apply filters
# -------------------------------
mask = pd.Series(True, index=filtered.index)

if selected_practice:
    mask &= filtered[practice_col].isin(selected_practice)

if selected_group_filter:
    mask &= filtered.apply(
        lambda row: selected_group in selected_group_filter if group_col is None else row[group_col] in selected_group_filter,
        axis=1
    )

if selected_cat_filter:
    mask &= filtered.apply(lambda row: sop_category in selected_cat_filter, axis=1)

if selected_region and len(selected_region) < len(region_present):
    region_mask = pd.Series(False, index=filtered.index)
    for r in selected_region:
        region_mask |= filtered["RegionsDetected"].str.contains(r, na=False)
    mask &= region_mask

filtered = filtered[mask].copy()

# -------------------------------
# Display filtered table
# -------------------------------
if filtered.empty:
    st.info(f"No SOPs found after applying filters.")
else:
    display_cols = []
    for c in [number_col, title_col, practice_col, group_col, bu_col, sop_type_col, "RegionsDetected", notes_col]:
        if c and c in filtered.columns and c not in display_cols:
            display_cols.append(c)

    table_df = filtered[display_cols].fillna("")
    rename_map = {
        number_col: "Number",
        title_col: "Title",
        practice_col: "Practice",
        group_col: "Group",
        bu_col: "Business Unit",
        sop_type_col: "SOP Type",
        notes_col: "Notes",
    }
    table_df = table_df.rename(columns={k: v for k, v in rename_map.items() if k})

    st.subheader(
        f"SOPs — Sheet: {sheet_choice} | Role: {selected_role_display} | Category: {sop_category}"
    )
    st.dataframe(table_df.reset_index(drop=True), use_container_width=True)

    # CSV download
    csv_buffer = io.StringIO()
    table_df.to_csv(csv_buffer, index=False)
    st.download_button(
        label="Download filtered SOPs as CSV",
        data=csv_buffer.getvalue().encode(),
        file_name=f"sops_filtered_{sheet_choice}.csv",
        mime="text/csv",
    )
