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

def checkbox_group_inline(values, key_prefix: str, default=True):
    """
    Returns list of values that are checked.
    Each checkbox key is deterministic and unique.
    """
    selected = []
    for i, v in enumerate(values):
        v_display = "(blank)" if (pd.isna(v) or str(v).strip() == "") else str(v)
        key = f"{key_prefix}__{i}"
        if key not in st.session_state:
            st.session_state[key] = default
        checked = st.checkbox(v_display, value=st.session_state[key], key=key)
        # keep session_state updated to reflect the current checked value
        st.session_state[key] = checked
        if checked:
            selected.append(v)
    return selected

# -------------------------------
# Load workbook (auto-select first sheet)
# -------------------------------
xls = pd.ExcelFile(excel_file_path)
sheets = xls.sheet_names
if not sheets:
    st.error("No sheets found in Excel file.")
    st.stop()

# AUTOMATIC: pick first sheet (removed "Choose sheet" UI)
sheet_choice = sheets[0]
st.markdown(f"**Using sheet:** {sheet_choice}")

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
# UI selections (main left area)
# -------------------------------
left_col, right_col = st.columns([3, 1])  # left: main content, right: filter pane

with left_col:
    selected_group = st.selectbox("Choose the Group (header group):", groups_sorted)

    role_entries = groups_map.get(selected_group, [])
    display_to_colidx = {}
    for col_idx, role_name in role_entries:
        display_to_colidx[f"{role_name} (col {col_idx})"] = col_idx

    selected_role_display = st.selectbox("Choose the Role (within group):", list(display_to_colidx.keys()))
    selected_col_idx = display_to_colidx[selected_role_display]

    # Keep a top convenience single-select SOP category (but sidebar multi-select will be authoritative)
    sop_category = st.selectbox("Convenience: SOP category (single-select)", list(category_map.keys()))
    st.markdown("---")

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

# Precompute role_series and RegionsDetected across entire data_df (so right pane can inspect)
role_series_all = data_df.iloc[:, selected_col_idx].apply(to_int_safe)
if notes_col:
    data_df["RegionsDetected"] = data_df[notes_col].apply(detect_regions)
else:
    data_df["RegionsDetected"] = ""

# -------------------------------
# RIGHT: Collapsible filter panel with Select All / Unselect All for Groups
# -------------------------------
with right_col:
    with st.expander("Filters (click to expand/collapse)", expanded=True):
        st.write("Use these checkboxes to filter results. Group list supports Select all / Unselect all.")

        # Practice filter
        practice_vals = []
        if practice_col and practice_col in data_df.columns:
            practice_vals = sorted(data_df[practice_col].dropna().unique())
            st.markdown("**Practice**")
            selected_practice = checkbox_group_inline(practice_vals, "filter_practice", default=True)
        else:
            selected_practice = None
            st.markdown("**Practice**")
            st.caption("No Practice column detected.")

        st.markdown("---")

        # Group filter (long list) with Select all / Unselect all controls
        group_vals = sorted(groups_sorted)
        st.markdown("**Group (header groups)**")
        # buttons that update session_state for each group checkbox key
        if st.button("Select all groups"):
            for i, g in enumerate(group_vals):
                key = f"filter_group__{i}"
                st.session_state[key] = True
        if st.button("Unselect all groups"):
            for i, g in enumerate(group_vals):
                key = f"filter_group__{i}"
                st.session_state[key] = False

        # Render group checkboxes (persistent keys)
        selected_group_filter = []
        for i, g in enumerate(group_vals):
            key = f"filter_group__{i}"
            if key not in st.session_state:
                st.session_state[key] = True  # default to checked
            checked = st.checkbox(g, value=st.session_state[key], key=key)
            st.session_state[key] = checked
            if checked:
                selected_group_filter.append(g)

        st.markdown("---")

        # Category filter (multi-checkbox)
        st.markdown("**Category**")
        cat_vals = list(category_map.keys())
        selected_cat_filter = checkbox_group_inline(cat_vals, "filter_cat", default=True)

        st.markdown("---")

        # Region filter (derived from notes)
        region_present = sorted({r for r in regions if any(data_df["RegionsDetected"].str.contains(r, na=False))})
        if region_present:
            st.markdown("**Region (from Notes)**")
            selected_region = checkbox_group_inline(region_present, "filter_region", default=True)
        else:
            selected_region = None
            st.markdown("**Region (from Notes)**")
            st.caption("No region indicators found in data.")

# -------------------------------
# Apply filters (combine into mask)
# -------------------------------
mask = pd.Series(True, index=data_df.index)

# Category filter: use sidebar multi-select (selected_cat_filter). Map to numeric codes.
if selected_cat_filter:
    allowed_codes = [category_map[c] for c in selected_cat_filter if c in category_map]
    if allowed_codes:
        mask &= role_series_all.isin(allowed_codes)
    else:
        # if no allowed_codes (user unchecked all), mask becomes all False
        mask &= False

# Practice
if selected_practice and practice_col and practice_col in data_df.columns:
    mask &= data_df[practice_col].isin(selected_practice)

# Group (data column) -- prefer filtering by data column if present; otherwise, filter by header groups selection
if group_col and group_col in data_df.columns:
    if selected_group_filter:
        # filter by values in the data 'Group' column
        mask &= data_df[group_col].isin(selected_group_filter)
else:
    # If user unchecked the header group that is currently selected in the main UI, no rows should remain.
    # Alternatively, we interpret header group filter as "keep the currently chosen header group only if it is checked".
    if selected_group not in selected_group_filter:
        mask &= False

# Region filter (OR across selected regions)
if selected_region is not None:
    if len(selected_region) < len(region_present):
        region_mask = pd.Series(False, index=data_df.index)
        for r in selected_region:
            region_mask |= data_df["RegionsDetected"].str.contains(r, na=False)
        mask &= region_mask
    # if all regions selected (default), no-op

filtered = data_df[mask].copy()

# -------------------------------
# Display filtered table (left column)
# -------------------------------
with left_col:
    if filtered.empty:
        st.info("No SOPs found after applying filters.")
    else:
        # Prepare display columns
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

        st.subheader(f"SOPs — Sheet: {sheet_choice} | Role: {selected_role_display}")
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
