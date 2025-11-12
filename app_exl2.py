# novotech_functional_training.py
import re
import io
import os
from typing import List, Set
import streamlit as st
import pandas as pd

st.set_page_config(page_title="Novotech Functional Training", layout="wide")
st.title("Novotech Functional Training")

# -------------------------------
# Path to Excel file (repo/data)
# -------------------------------
EXCEL_PATH = os.path.join(os.path.dirname(__file__), "data", "PR_Audience_LearningItem_Export.xlsx")
if not os.path.exists(EXCEL_PATH):
    st.error(f"Expected Excel file not found at {EXCEL_PATH}")
    st.stop()
st.write(f"Reading Excel from: {EXCEL_PATH}")

# -------------------------
# Parsing helpers for Column B
# -------------------------
ROLE_RE = re.compile(r"Role\s*:\s*Any of\s*:\s*\(([^)]*)\)", flags=re.I)
ORG_RE = re.compile(r"Organisation\s*:\s*Any of\s*:\s*\(([^)]*)\)", flags=re.I)

def split_items(s: str) -> List[str]:
    if not isinstance(s, str) or s.strip() == "":
        return []
    parts = [p.strip() for p in re.split(r",\s*", s) if p is not None]
    parts = [p for p in parts if p != ""]
    return parts

def extract_roles(cell: str) -> List[str]:
    if not isinstance(cell, str):
        return []
    roles = []
    for m in ROLE_RE.finditer(cell):
        roles += split_items(m.group(1))
    roles = [re.sub(r"^[\(\)\s]+|[\(\)\s]+$", "", r) for r in roles]
    return sorted(set([r for r in roles if r]))

def extract_org_groups_practices(cell: str):
    if not isinstance(cell, str):
        return [], []
    groups = []
    practices = []
    for m in ORG_RE.finditer(cell):
        items = split_items(m.group(1))
        for it in items:
            it = it.strip()
            g_match = re.match(r"^(.*?)(?:\s*\(\s*Group\s*\)\s*)$", it, flags=re.I)
            p_match = re.match(r"^(.*?)(?:\s*\(\s*Practice\s*\)\s*)$", it, flags=re.I)
            if g_match:
                groups.append(g_match.group(1).strip())
            elif p_match:
                practices.append(p_match.group(1).strip())
            else:
                # ambiguous entries without suffix ignored
                pass
    return sorted(set([g for g in groups if g])), sorted(set([p for p in practices if p]))

# -------------------------
# Load Excel and columns A..E with header-row cleanup
# -------------------------
xls = pd.ExcelFile(EXCEL_PATH)
sheet_name = xls.sheet_names[0]
raw = pd.read_excel(EXCEL_PATH, sheet_name=sheet_name, header=None, dtype=object)

# require at least 5 columns (A-E)
if raw.shape[1] < 5:
    st.error(f"Expected at least 5 columns (A-E). Found {raw.shape[1]}.")
    st.stop()

# Detect and drop header-like rows (first few rows)
header_tokens = {"prescriptive", "prescriptive rule", "member selection", "course id", "course title", "curriculum"}
rows_to_drop = []
for idx in range(min(6, raw.shape[0])):
    row_vals = " ".join([str(x).lower() for x in raw.iloc[idx, :5].tolist()])
    for tok in header_tokens:
        if tok in row_vals:
            rows_to_drop.append(idx)
            break

if rows_to_drop:
    raw = raw.drop(rows_to_drop).reset_index(drop=True)
    st.write(f"Dropped header-like row(s): {rows_to_drop}")

# extract columns A..E (0..4)
prescriptive = raw.iloc[:, 0].astype(object).fillna("").astype(str)
member_criteria = raw.iloc[:, 1].astype(object).fillna("").astype(str)
course_id_col = raw.iloc[:, 2].astype(object).fillna("").astype(str)
course_title_col = raw.iloc[:, 3].astype(object).fillna("").astype(str)
curriculum_title_col = raw.iloc[:, 4].astype(object).fillna("").astype(str)

nrows = len(raw)
st.write(f"Loaded {nrows} data rows (after header cleanup).")

# -------------------------
# Group rows by prescriptive rule (global duplicates -> curriculum)
# -------------------------
rule_to_indices = {}
for idx, rule in enumerate(prescriptive):
    key = rule.strip()
    rule_to_indices.setdefault(key, []).append(idx)

records = []

def parse_row_attributes(index: int):
    cell = member_criteria[index]
    roles = extract_roles(cell)
    groups, practices = extract_org_groups_practices(cell)
    return groups, practices, roles

for rule_key, indices in rule_to_indices.items():
    if rule_key != "" and len(indices) > 1:
        # curriculum record
        curr_title = ""
        for i in indices:
            t = curriculum_title_col[i]
            if isinstance(t, str) and t.strip() != "":
                curr_title = t.strip()
                break
        if not curr_title:
            curr_title = rule_key
        agg_groups: Set[str] = set()
        agg_practices: Set[str] = set()
        agg_roles: Set[str] = set()
        for i in indices:
            g, p, r = parse_row_attributes(i)
            agg_groups.update(g)
            agg_practices.update(p)
            agg_roles.update(r)
        records.append({
            "Title": curr_title,
            "ID": "",
            "Type": "curriculum",
            "Groups": sorted(list(agg_groups)),
            "Practices": sorted(list(agg_practices)),
            "Roles": sorted(list(agg_roles)),
            "SourceIndices": indices
        })
        # add course rows
        for i in indices:
            title = course_title_col[i].strip() if isinstance(course_title_col[i], str) else ""
            cid = course_id_col[i].strip() if isinstance(course_id_col[i], str) else ""
            g, p, r = parse_row_attributes(i)
            records.append({
                "Title": title if title else rule_key,
                "ID": cid,
                "Type": "course",
                "Groups": sorted(list(g)),
                "Practices": sorted(list(p)),
                "Roles": sorted(list(r)),
                "SourceIndex": i
            })
    else:
        # single occurrence -> course
        i = indices[0]
        title = course_title_col[i].strip() if isinstance(course_title_col[i], str) else ""
        cid = course_id_col[i].strip() if isinstance(course_id_col[i], str) else ""
        g, p, r = parse_row_attributes(i)
        records.append({
            "Title": title if title else (curriculum_title_col[i].strip() if isinstance(curriculum_title_col[i], str) and curriculum_title_col[i].strip() else prescriptive[i]),
            "ID": cid,
            "Type": "course",
            "Groups": sorted(list(g)),
            "Practices": sorted(list(p)),
            "Roles": sorted(list(r)),
            "SourceIndex": i
        })

# -------------------------
# DataFrame of parsed records
# -------------------------
out_df = pd.DataFrame(records)

for col in ["Groups", "Practices", "Roles"]:
    if col not in out_df.columns:
        out_df[col] = [[] for _ in range(len(out_df))]
    else:
        out_df[col] = out_df[col].apply(lambda v: v if isinstance(v, list) else ([] if pd.isna(v) else [v]))

# -------------------------
# Filter lists (alphabetical)
# -------------------------
all_practices = sorted({p for row in out_df["Practices"] for p in row}, key=lambda x: x.lower())
all_groups = sorted({g for row in out_df["Groups"] for g in row}, key=lambda x: x.lower())
all_roles = sorted({r for row in out_df["Roles"] for r in row}, key=lambda x: x.lower())

# -------------------------
# UI layout: left filters, right table
# -------------------------
left_col, right_col = st.columns([1, 3])

def checkbox_list_unchecked(label: str, options: List[str], key_prefix: str) -> List[str]:
    st.markdown(f"**{label}**")
    selected = []
    for i, opt in enumerate(options):
        key = f"{key_prefix}__{i}"
        checked = st.checkbox(opt, value=False, key=key)
        if checked:
            selected.append(opt)
    return selected

with left_col:
    st.header("Filters")
    selected_practices = checkbox_list_unchecked("Practice", all_practices, "flt_practice")
    st.markdown("---")
    selected_groups = checkbox_list_unchecked("Group", all_groups, "flt_group")
    st.markdown("---")
    selected_roles = checkbox_list_unchecked("Role", all_roles, "flt_role")
    st.markdown("---")
    st.write("Leave all checkboxes unchecked to include all values for that filter.")
    st.write("Logic: OR within a filter, AND across filters.")

# -------------------------
# Apply filters
# -------------------------
def row_matches_filter_list(row_list: List[str], selected_list: List[str]) -> bool:
    if not selected_list:
        return True
    if not row_list:
        return False
    return any(item in selected_list for item in row_list)

mask = out_df.apply(
    lambda r: (
        row_matches_filter_list(r["Practices"], selected_practices) and
        row_matches_filter_list(r["Groups"], selected_groups) and
        row_matches_filter_list(r["Roles"], selected_roles)
    ),
    axis=1
)
filtered_df = out_df[mask].copy()

# -------------------------
# Display results (Title, ID, Type) and CSV download
# -------------------------
with right_col:
    st.subheader("Filtered Learning Items")
    if filtered_df.empty:
        st.info("No items match the current filters.")
    else:
        display_df = filtered_df[["Title", "ID", "Type"]].copy()
        display_df["ID"] = display_df["ID"].replace("", "")
        st.dataframe(display_df.reset_index(drop=True), use_container_width=True)
        csv_buf = io.StringIO()
        display_df.to_csv(csv_buf, index=False)
        st.download_button(
            label="Download visible items as CSV",
            data=csv_buf.getvalue().encode(),
            file_name="novotech_functional_training_filtered.csv",
            mime="text/csv"
        )

# -------------------------
# Summary
# -------------------------
st.markdown("---")
st.write(f"Total parsed items: {len(out_df)}. Showing {len(filtered_df)} after filters.")
if st.checkbox("Show parsed items with attributes (debug)", value=False):
    debug_df = out_df.copy()
    for c in ["Groups", "Practices", "Roles"]:
        debug_df[c] = debug_df[c].apply(lambda L: ", ".join(L) if isinstance(L, list) else "")
    st.dataframe(debug_df.reset_index(drop=True), use_container_width=True)
