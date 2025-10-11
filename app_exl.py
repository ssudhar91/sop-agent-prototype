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
# Settings
# -------------------------------
# The user said Row 3 contains headers ("Business Unit","SOP Type","Number","Title" and roles from E3 onwards).
HEADER_ROW = 2  # zero-indexed: row 3 in Excel

category_map = {
    "Within 2 weeks": 1,    # All staff SOPs
    "Within 90 days": 2,    # Role-based SOPs
    "Before task": 3        # Optional: before a particular task
}

# -------------------------------
# Utility: robust column lookup
# -------------------------------
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

# -------------------------------
# Read sheets and extract roles
# -------------------------------
xls = pd.ExcelFile(excel_file_path)
# Read the first sheet header to get role column names (we assume same structure across sheets)
first_sheet = xls.sheet_names[0]
df_head = pd.read_excel(excel_file_path, sheet_name=first_sheet, header=HEADER_ROW)

# Roles start from column E (index 4)
if len(df_head.columns) <= 4:
    st.error("Unable to detect role columns: sheet doesn't have columns beyond E.")
    st.stop()

role_columns = list(df_head.columns[4:])
roles = role_columns.copy()  # Use actual header names as role labels for dropdown

# -------------------------------
# User selections
# -------------------------------
sop_category = st.selectbox("Choose the SOP category:", list(category_map.keys()))
selected_role = st.selectbox("Choose your role:", roles)

st.markdown("---")

# -------------------------------
# Function to process a sheet into a normalized dataframe
# -------------------------------
def process_sheet(sheet_name):
    df = pd.read_excel(excel_file_path, sheet_name=sheet_name, header=HEADER_ROW, dtype=object)
    # Ensure role columns exist; if not, skip
    if len(df.columns) <= 4:
        return pd.DataFrame()  # empty

    # Normalize column names we expect
    bu_col = pick_column(df, ["Business Unit", "BusinessUnit", "Business unit", "Business_unit"])
    sop_type_col = pick_column(df, ["SOP Type", "SOPType", "SOP type"])
    number_col = pick_column(df, ["Number", "No", "ID", "SOP Number", "Number "])
    title_col = pick_column(df, ["Title", "SOP Title", "Name", "Title "])
    notes_col = pick_column(df, ["Notes", "Note", "Remarks", "Region Notes", "Comments"])

    # If title or number missing try to find likely alternate names
    if title_col is None:
        # take the first unused non-role column after D
        for c in df.columns[:4]:
            # ignore the first 4 header columns if they look like the standard ones
            pass
        # fallback to column D if exists
        if len(df.columns) >= 4:
            title_col = df.columns[3]

    # Prepare a consistent dataframe
    # Keep original role columns (E onwards) as-is
    role_cols = list(df.columns[4:])
    normalized = pd.DataFrame()
    normalized["__sheet__"] = sheet_name
    normalized["Business Unit"] = df[bu_col] if bu_col in df.columns else ""
    normalized["SOP Type"] = df[sop_type_col] if sop_type_col in df.columns else ""
    normalized["Number"] = df[number_col] if number_col in df.columns else ""
    normalized["Title"] = df[title_col] if title_col in df.columns else ""
    normalized["Notes"] = df[notes_col] if (notes_col and notes_col in df.columns) else ""

    # attach role columns as they are (keep original values)
    for rc in role_cols:
        normalized[rc] = df[rc]

    return normalized

# -------------------------------
# Aggregate all sheets
# -------------------------------
all_sops = []
for sheet_name in xls.sheet_names:
    proc = process_sheet(sheet_name)
    if not proc.empty:
        all_sops.append(proc)

if not all_sops:
    st.warning("No structured SOP rows found across sheets.")
    st.stop()

master_df = pd.concat(all_sops, ignore_index=True)

# Convert role columns values to numeric when possible (strip spaces)
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

# Apply conversion for role columns
for rc in roles:
    master_df[rc] = master_df[rc].apply(to_int_safe)

# -------------------------------
# Filter SOPs by selected role + category
# -------------------------------
category_value = category_map[sop_category]

filtered = master_df[master_df[selected_role] == category_value].copy()

# -------------------------------
# Prepare display table
# -------------------------------
if filtered.empty:
    st.info(f"No SOPs found for role **{selected_role}** in category **{sop_category}**.")
else:
    # Build a neat table with required columns: Number, Title, Business Unit, SOP Type, Notes, Sheet
    display_cols = ["Number", "Title", "Business Unit", "SOP Type", "Notes", "__sheet__"]
    # If any of these columns are missing in the DF, only keep ones that exist
    display_cols = [c for c in display_cols if c in filtered.columns]

    table_df = filtered[display_cols].copy()
    # Clean up NaNs
    table_df = table_df.fillna("")

    # Rename __sheet__ to Sheet for display
    if "__sheet__" in table_df.columns:
        table_df = table_df.rename(columns={"__sheet__": "Sheet"})

    # Ensure Number and Title are first columns (Number then Title)
    cols_order = []
    if "Number" in table_df.columns:
        cols_order.append("Number")
    if "Title" in table_df.columns:
        cols_order.append("Title")
    # then others
    for c in table_df.columns:
        if c not in cols_order:
            cols_order.append(c)
    table_df = table_df[cols_order]

    st.subheader(f"SOPs for role **{selected_role}** — Category: **{sop_category}**")
    st.dataframe(table_df, use_container_width=True)

    # Provide CSV download
    csv_buffer = io.StringIO()
    table_df.to_csv(csv_buffer, index=False)
    csv_bytes = csv_buffer.getvalue().encode()
    st.download_button(
        label="Download filtered SOPs as CSV",
        data=csv_bytes,
        file_name=f"sops_{selected_role}_{sop_category.replace(' ', '_')}.csv",
        mime="text/csv",
    )

# -------------------------------
# Extra: show region-specific SOPs if present in Notes
# -------------------------------
st.markdown("---")
st.write("Region-specific SOPs (detected in Notes):")
regions = ["china", "korea", "taiwan", "hong kong", "india", "us", "uk"]  # add more as desired
def detect_regions(text):
    text = str(text).lower()
    found = [r for r in regions if r in text]
    return ", ".join(found) if found else ""

master_df["RegionsDetected"] = master_df["Notes"].apply(detect_regions) if "Notes" in master_df.columns else ""
region_hits = master_df[master_df["RegionsDetected"] != ""]
if not region_hits.empty:
    st.dataframe(region_hits[["Number", "Title", "Business Unit", "RegionsDetected"]].fillna(""), use_container_width=True)
else:
    st.write("No region-specific information detected (based on common region keywords).")
