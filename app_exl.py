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

st.write(f"Using master Excel file: {excel_file_path}")

# -------------------------------
# Settings: header rows (0-indexed)
# Based on your description:
# - Group names row is HEADER_GROUP_ROW
# - Column headers (Business Unit, SOP Type, Number, Title, Roles...) are in HEADER_COLS_ROW
# - Data starts at DATA_START_ROW
HEADER_GROUP_ROW = 0
HEADER_COLS_ROW = 2
DATA_START_ROW = 3  # zero-indexed (Excel row 4)

category_map = {
    "Within 2 weeks": 1,    # All staff SOPs
    "Within 90 days": 2,    # Role-based SOPs
    "Before task": 3        # Optional: before a particular task
}

# -------------------------------
# Helper functions
def pick_column(df, candidates):
    """Return the first matching column name from candidates (case-insensitive)."""
    cols_lower = {c.lower(): c for c in df.columns}
    for cand in candidates:
        if cand is None:
            continue
        key = cand.lower()
        if key in cols_lower:
            return cols_lower[key]
    return None

def to_int_safe(v):
    try:
        if pd.isna(v):
            return None
        if isinstance(v, str):
            v2 = v.strip()
            if v2 == "":
                return None
            return int(float(v2))
        return int(v)
    except Exception:
        return None

# -------------------------------
# Load workbook and list sheets
xls = pd.ExcelFile(excel_file_path)
sheets = xls.sheet_names
sheet_choice = st.selectbox("Choose sheet:", sheets)

# Load the selected sheet with no header to access the group row and header row
raw = pd.read_excel(excel_file_path, sheet_name=sheet_choice, header=None, dtype=object)

# Basic sanity checks
if raw.shape[0] <= HEADER_COLS_ROW:
    st.error("The selected sheet doesn't have the expected header rows. Check the sheet layout.")
    st.stop()

# Extract group row and header (roles) row
group_row = raw.iloc[HEADER_GROUP_ROW].copy().fillna(method="ffill")
header_row = raw.iloc[HEADER_COLS_ROW].astype(str).tolist()

# Build data frame with header_row as columns and data starting from DATA_START_ROW
data_df = raw.iloc[DATA_START_ROW:].copy().reset_index(drop=True)
data_df.columns = header_row

# Roles are columns from index 4 (E) onward per your spec
if len(header_row) <= 4:
    st.error("Unable to detect role columns: sheet doesn't have columns beyond column E.")
    st.stop()

# Build group -> list of (col_idx, role_name) mapping using positions to avoid duplicate-label issues
groups_map = {}
for col_idx, col_name in enumerate(header_row):
    if col_idx < 4:
        continue
    raw_group_val = group_row.iloc[col_idx] if col_idx < len(group_row) else ""
    group_name = str(raw_group_val).strip() if not pd.isna(raw_group_val) else ""
    if group_name == "" or group_name.lower() in ("nan", "none"):
        group_name = "Ungrouped"
    groups_map.setdefault(group_name, []).append((col_idx, col_name))

# Sort groups for UI
groups_sorted = sorted(groups_map.keys())

# -------------------------------
# UI: choose group then role
selected_group = st.selectbox("Choose the Group:", groups_sorted)

# Build display options for roles within the selected group.
# To avoid ambiguity we map display string -> column index.
role_entries = groups_map.get(selected_group, [])
display_to_colidx = {}
display_options = []
for i, (col_idx, role_name) in enumerate(role_entries):
    # Make a readable label. Include column index to ensure uniqueness (in case role repeats within group).
    display_label = f"{role_name}  (col {col_idx})"
    display_options.append(display_label)
    display_to_colidx[display_label] = col_idx

selected_role_display = st.selectbox("Choose the Role (within selected group):", display_options)
selected_col_idx = display_to_colidx[selected_role_display]

# -------------------------------
# Choose SOP Category
sop_category = st.selectbox("Choose the SOP category:", list(category_map.keys()))
category_value = category_map[sop_category]

st.markdown("---")

# -------------------------------
# Normalize expected columns (Number, Title, Business Unit, SOP Type, Notes)
# Use header names present in data_df (which are header_row entries)
bu_col = pick_column(data_df, ["Business Unit", "BusinessUnit", "Business unit", "Business_unit"])
sop_type_col = pick_column(data_df, ["SOP Type", "SOPType", "SOP type"])
number_col = pick_column(data_df, ["Number", "No", "ID", "SOP Number", "SOP No", "Number "])
title_col = pick_column(data_df, ["Title", "SOP Title", "Name", "Title "])
notes_col = pick_column(data_df, ["Notes", "Note", "Remarks", "Region Notes", "Comments"])

# Fallback: if title_col missing, try column D (4th column name)
if title_col is None and len(data_df.columns) >= 4:
    title_col = data_df.columns[3]

# -------------------------------
# Use iloc with selected_col_idx to get the exact role column (avoids duplicate label problem)
role_series = data_df.iloc[:, selected_col_idx].apply(to_int_safe)

# Filter rows where role == category_value
filtered = data_df[role_series == category_value].copy()

# -------------------------------
# Prepare table to display with Number & Title first
if filtered.empty:
    st.info(f"No SOPs found for role **{selected_role_display}** in category **{sop_category}** (sheet: {sheet_choice}).")
else:
    # Build display dataframe
    display_cols = []
    if number_col and number_col in filtered.columns:
        display_cols.append(number_col)
    if title_col and title_col in filtered.columns:
        display_cols.append(title_col)

    # Add additional useful columns
    for c in ["Business Unit", "SOP Type", "Notes"]:
        if c in filtered.columns:
            display_cols.append(c)

    # If none of number/title detected, fall back to first 4 columns
    if not display_cols:
        display_cols = list(filtered.columns[:4])

    table_df = filtered[display_cols].copy()
    # Clean NaNs
    table_df = table_df.fillna("")
    # Standardize column names for display: prefer "Number" and "Title"
    rename_map = {}
    if number_col and number_col in table_df.columns:
        rename_map[number_col] = "Number"
    if title_col and title_col in table_df.columns:
        rename_map[title_col] = "Title"
    table_df = table_df.rename(columns=rename_map)

    # Reorder to ensure Number then Title
    cols_order = []
    if "Number" in table_df.columns:
        cols_order.append("Number")
    if "Title" in table_df.columns:
        cols_order.append("Title")
    for c in table_df.columns:
        if c not in cols_order:
            cols_order.append(c)
    table_df = table_df[cols_order]

    st.subheader(
        f"SOPs — Sheet: **{sheet_choice}** | Group: **{selected_group}** | Role: **{selected_role_display}** | Category: **{sop_category}**"
    )
    st.dataframe(table_df.reset_index(drop=True), use_container_width=True)

    # CSV download
    csv_buffer = io.StringIO()
    table_df.to_csv(csv_buffer, index=False)
    csv_bytes = csv_buffer.getvalue().encode()
    st.download_button(
        label="Download filtered SOPs as CSV",
        data=csv_bytes,
        file_name=f"sops_{sheet_choice}_{selected_group}_{selected_role_display}_{sop_category.replace(' ','_')}.csv",
        mime="text/csv",
    )

# -------------------------------
# Region detection (from Notes) — optional
st.markdown("---")
st.write("Region-specific SOPs detected in Notes (simple keyword scan):")
regions = ["china", "korea", "taiwan", "hong kong", "india", "us", "uk"]  # extend as needed

if notes_col and notes_col in data_df.columns:
    def detect_regions(text):
        text = str(text).lower()
        found = [r for r in regions if r in text]
        return ", ".join(found) if found else ""

    data_df["RegionsDetected"] = data_df[notes_col].apply(detect_regions)
    region_hits = data_df[data_df["RegionsDetected"] != ""]
    if not region_hits.empty:
        # prepare display
        display = region_hits[[c for c in [number_col, title_col, "RegionsDetected"] if c in region_hits.columns]].fillna("")
        rename_map2 = {}
        if number_col and number_col in display.columns:
            rename_map2[number_col] = "Number"
        if title_col and title_col in display.columns:
            rename_map2[title_col] = "Title"
        display = display.rename(columns=rename_map2)
        st.dataframe(display.reset_index(drop=True), use_container_width=True)
    else:
        st.write("No region-specific indicators found.")
else:
    st.write("No Notes column to detect region-specific SOPs.")
