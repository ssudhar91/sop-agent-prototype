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
# -------------------------------
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
# -------------------------------
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
# -------------------------------
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
# -------------------------------
sop_category = st.selectbox("Choose the SOP category:", list(category_map.keys()))
category_value = category_map[sop_category]

st.markdown("---")

# -------------------------------
# Normalize expected columns (Number, Title, Business Unit, SOP Type, Notes)
# -------------------------------
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
# -------------------------------
role_series = data_df.iloc[:, selected_col_idx].apply(to_int_safe)

# Filter rows where role == category_value
filtered = data_df[role_series == category_value].copy()

# =====================================================
# NEW: Sidebar filter pane with checkboxes
# =====================================================

# Prepare a RegionsDetected column *for filtered set* to enable region filtering.
regions = ["china", "korea", "taiwan", "hong kong", "india", "us", "uk"]  # extend as needed

def detect_regions(text):
    text = str(text).lower()
    found = [r for r in regions if r in text]
    return ", ".join(found) if found else ""

if notes_col and notes_col in filtered.columns:
    filtered["RegionsDetected"] = filtered[notes_col].apply(detect_regions)
else:
    filtered["RegionsDetected"] = ""

def checkbox_group(label: str, values, key_prefix: str):
    """
    Render a group of checkboxes and return the selected values.
    By default, all options are checked.
    """
    st.markdown(f"**{label}**")
    selected = []
    for v in values:
        vs = "" if (pd.isna(v) or v is None) else str(v)
        # stable, unique key per value
        if st.checkbox(vs if vs != "" else "(blank)", value=True, key=f"{key_prefix}_{vs}"):
            selected.append(v)
    return selected

with st.sidebar:
    st.header("Filters")

    # Business Unit filter
    if bu_col and bu_col in filtered.columns:
        bu_vals = list(pd.Series(filtered[bu_col].astype(str)).replace("nan", "").unique())
        bu_vals = sorted(bu_vals, key=lambda x: (x == "", x.lower()))
        selected_bu = checkbox_group("Business Unit", bu_vals, "bu")
    else:
        selected_bu = None

    # SOP Type filter
    if sop_type_col and sop_type_col in filtered.columns:
        st.markdown("---")
        sop_vals = list(pd.Series(filtered[sop_type_col].astype(str)).replace("nan", "").unique())
        sop_vals = sorted(sop_vals, key=lambda x: (x == "", x.lower()))
        selected_sop = checkbox_group("SOP Type", sop_vals, "soptype")
    else:
        selected_sop = None

    # Region filter (derived from Notes keyword scan)
    st.markdown("---")
    # Offer checkboxes only for regions that actually occur in the filtered set
    region_present = set()
    if "RegionsDetected" in filtered.columns:
        for r in regions:
            if any(filtered["RegionsDetected"].str.contains(r, na=False)):
                region_present.add(r)
    region_present = sorted(region_present)
    if region_present:
        selected_regions = checkbox_group("Regions (from Notes)", region_present, "region")
    else:
        selected_regions = None
        st.caption("No region indicators detected in current selection.")

# Apply filters to 'filtered'
mask = pd.Series(True, index=filtered.index)

# Business Unit mask
if selected_bu is not None:
    # treat blanks explicitly
    bu_series = filtered[bu_col].astype(str).replace("nan", "")
    mask &= bu_series.isin(selected_bu)

# SOP Type mask
if selected_sop is not None:
    sop_series = filtered[sop_type_col].astype(str).replace("nan", "")
    mask &= sop_series.isin(selected_sop)

# Regions mask (OR across multiple regions; rows with no region tag are excluded if any region is selected)
if selected_regions is not None and len(selected_regions) > 0:
    if len(selected_regions) != len(region_present):  # only filter if user unticked something
        region_mask = pd.Series(False, index=filtered.index)
        for r in selected_regions:
            region_mask |= filtered["RegionsDetected"].str.contains(r, na=False)
        mask &= region_mask

filtered = filtered[mask].copy()
# =====================================================
# END NEW: Sidebar filter pane
# =====================================================

# -------------------------------
# Prepare table to display with Number & Title first
# -------------------------------
if filtered.empty:
    st.info(
        f"No SOPs found for role **{selected_role_display}** in category **{sop_category}** "
        f"(sheet: {sheet_choice}) after applying filters."
    )
else:
    # Build display dataframe
    display_cols = []
    if number_col and number_col in filtered.columns:
        display_cols.append(number_col)
    if title_col and title_col in filtered.columns:
        display_cols.append(title_col)

    # Add additional useful columns
    for c in ["Business Unit", "SOP Type", "Notes", "RegionsDetected"]:
        if c in filtered.columns:
            display_cols.append(c)
        else:
            # If our normalized names differ, try the resolved names
            if c == "Business Unit" and bu_col and bu_col in filtered.columns:
                display_cols.append(bu_col)
            if c == "SOP Type" and sop_type_col and sop_type_col in filtered.columns:
                display_cols.append(sop_type_col)
            if c == "Notes" and notes_col and notes_col in filtered.columns:
                display_cols.append(notes_col)

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
    if bu_col and bu_col in table_df.columns:
        rename_map[bu_col] = "Business Unit"
    if sop_type_col and sop_type_col in table_df.columns:
        rename_map[sop_type_col] = "SOP Type"
    if notes_col and notes_col in table_df.columns:
        rename_map[notes_col] = "Notes"
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
# Region detection (from Notes) — optional (summary view)
# -------------------------------
st.markdown("---")
st.write("Region-specific SOPs detected in Notes (simple keyword scan):")
# Reuse the table already computed (filtered has RegionsDetected)
region_hits = filtered[filtered["RegionsDetected"] != ""] if "RegionsDetected" in filtered.columns else pd.DataFrame()

if not region_hits.empty:
    display_cols2 = []
    if number_col and number_col in region_hits.columns:
        display_cols2.append(number_col)
    if title_col and title_col in region_hits.columns:
        display_cols2.append(title_col)
    display_cols2.append("RegionsDetected")
    display2 = region_hits[display_cols2].fillna("")
    rename_map2 = {}
    if number_col and number_col in display2.columns:
        rename_map2[number_col] = "Number"
    if title_col and title_col in display2.columns:
        rename_map2[title_col] = "Title"
    display2 = display2.rename(columns=rename_map2)
    st.dataframe(display2.reset_index(drop=True), use_container_width=True)
else:
    st.write("No region-specific indicators found (after current filters).")
