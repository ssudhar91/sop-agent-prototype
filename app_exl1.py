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
# Path to master Excel file (same as before)
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
# Helpers
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
            v2 = v.strip()
            if v2 == "":
                return None
            return int(float(v2))
        return int(v)
    except Exception:
        return None

def detect_regions(text):
    text = str(text).lower()
    found = [r for r in regions if r in text]
    return ", ".join(found) if found else ""

# -------------------------------
# Load workbook - automatic first sheet (no sheet chooser)
# -------------------------------
xls = pd.ExcelFile(excel_file_path)
sheets = xls.sheet_names
if not sheets:
    st.error("No sheets found in the Excel file.")
    st.stop()

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

# Build group -> list of (col_idx, role_name)
groups_map = {}
for col_idx, col_name in enumerate(header_row):
    if col_idx < 4:
        continue
    raw_group_val = group_row.iloc[col_idx] if col_idx < len(group_row) else ""
    group_name = str(raw_group_val).strip() if not pd.isna(raw_group_val) else ""
    if group_name == "" or group_name.lower() in ("nan", "none"):
        group_name = "Ungrouped"
    groups_map.setdefault(group_name, []).append((col_idx, col_name))

# Sort group names alphabetically
groups_sorted = sorted(groups_map.keys(), key=lambda x: x.lower())

# Identify common columns
practice_col = pick_column(data_df, ["Practice", "Department", "Function"])
group_col = pick_column(data_df, ["Group", "SOP Group", "Team Group"])
bu_col = pick_column(data_df, ["Business Unit", "BusinessUnit"])
sop_type_col = pick_column(data_df, ["SOP Type", "Type"])
number_col = pick_column(data_df, ["Number", "SOP Number", "No", "ID"])
title_col = pick_column(data_df, ["Title", "SOP Title"])
notes_col = pick_column(data_df, ["Notes", "Remarks", "Comments", "Region Notes"])

if title_col is None and len(data_df.columns) >= 4:
    title_col = data_df.columns[3]

# Compute RegionsDetected for rows (useful if needed later)
if notes_col:
    data_df["RegionsDetected"] = data_df[notes_col].apply(detect_regions)
else:
    data_df["RegionsDetected"] = ""

# -------------------------------
# Layout: left = filters, right = main (role selection + results)
# -------------------------------
left_col, right_col = st.columns([1, 3])

# LEFT: Filters (Group & Category) - unchecked by default
with left_col:
    st.header("Filters")

    st.markdown("**Group** (header groups) — unchecked by default")
    # Build checkbox list (unchecked)
    selected_groups = []
    for i, g in enumerate(groups_sorted):
        key = f"filter_group__{i}"
        # default False for unchecked initially
        checked = st.checkbox(g, value=False, key=key)
        if checked:
            selected_groups.append(g)

    st.markdown("---")
    st.markdown("**Category** — unchecked by default")
    selected_categories = []
    cat_keys = list(category_map.keys())
    for i, c in enumerate(cat_keys):
        key = f"filter_cat__{i}"
        checked = st.checkbox(c, value=False, key=key)
        if checked:
            selected_categories.append(c)

    st.markdown("---")
    st.write("Tip: leave all checkboxes unchecked to show all Groups / Categories.")

# RIGHT: role selector + results
with right_col:
    # Role selection remains available; but results are driven by left filters
    # If no group checked, show all groups in the role dropdown; otherwise show only checked groups
    groups_for_role_select = selected_groups if selected_groups else groups_sorted

    # If there are no groups (edge case) stop
    if not groups_for_role_select:
        st.error("No header groups detected.")
        st.stop()

    selected_group_for_role = st.selectbox("Choose the Group (for role list):", groups_for_role_select)
    # Build display options for roles for that group
    role_entries = groups_map.get(selected_group_for_role, [])
    display_to_colidx = {}
    role_display_options = []
    for col_idx, role_name in role_entries:
        label = f"{role_name} (col {col_idx})"
        display_to_colidx[label] = col_idx
        role_display_options.append(label)

    if not role_display_options:
        st.warning("No roles found under the selected header group.")
        selected_role_display = None
        selected_col_idx = None
    else:
        selected_role_display = st.selectbox("Choose the Role (within group):", role_display_options)
        selected_col_idx = display_to_colidx[selected_role_display]

    st.markdown("---")

    # Build filter logic:
    # - Groups: if none selected -> all groups_allowed = groups_sorted
    #           else groups_allowed = selected_groups (OR across groups)
    # - Categories: if none selected -> all codes allowed
    #               else allowed_codes = [category_map[name] ...] (OR across categories)
    groups_allowed = selected_groups if selected_groups else groups_sorted
    # collect allowed column indices for those groups
    allowed_cols = []
    for g in groups_allowed:
        allowed_cols += [col_idx for col_idx, _ in groups_map.get(g, [])]
    allowed_cols = sorted(set(allowed_cols))

    # categories logic
    if selected_categories:
        allowed_codes = [category_map[c] for c in selected_categories if c in category_map]
    else:
        # allow all categories if none selected
        allowed_codes = list(category_map.values())

    # Now build mask: include a row if ANY of the allowed_cols has a value in allowed_codes
    if not allowed_cols:
        # no role columns found for allowed groups -> no results
        mask = pd.Series(False, index=data_df.index)
    else:
        # gather numeric values for allowed columns safely
        # create DataFrame of numeric parsed values for allowed cols
        allowed_vals_df = data_df.iloc[:, allowed_cols].applymap(to_int_safe)
        # mask True if any column in row has a value in allowed_codes
        mask = allowed_vals_df.isin(allowed_codes).any(axis=1)

    filtered = data_df[mask].copy()

    # Prepare display columns (Number then Title first)
    if filtered.empty:
        st.info("No SOPs found for the current Group(s)/Category selection.")
    else:
        display_cols = []
        if number_col and number_col in filtered.columns:
            display_cols.append(number_col)
        if title_col and title_col in filtered.columns:
            display_cols.append(title_col)

        # add commonly useful columns if present
        for c in [practice_col, group_col, bu_col, sop_type_col, "RegionsDetected", notes_col]:
            if c and c in filtered.columns and c not in display_cols:
                display_cols.append(c)

        if not display_cols:
            display_cols = list(filtered.columns[:6])

        table_df = filtered[display_cols].fillna("")

        # rename for nicer display
        rename_map = {}
        if number_col and number_col in table_df.columns:
            rename_map[number_col] = "Number"
        if title_col and title_col in table_df.columns:
            rename_map[title_col] = "Title"
        if practice_col and practice_col in table_df.columns:
            rename_map[practice_col] = "Practice"
        if group_col and group_col in table_df.columns:
            rename_map[group_col] = "Group"
        if bu_col and bu_col in table_df.columns:
            rename_map[bu_col] = "Business Unit"
        if sop_type_col and sop_type_col in table_df.columns:
            rename_map[sop_type_col] = "SOP Type"
        if notes_col and notes_col in table_df.columns:
            rename_map[notes_col] = "Notes"

        table_df = table_df.rename(columns=rename_map)

        # ensure Number then Title order
        cols_order = []
        if "Number" in table_df.columns:
            cols_order.append("Number")
        if "Title" in table_df.columns:
            cols_order.append("Title")
        for c in table_df.columns:
            if c not in cols_order:
                cols_order.append(c)
        table_df = table_df[cols_order]

        st.subheader(f"SOPs — Sheet: {sheet_choice}")
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
