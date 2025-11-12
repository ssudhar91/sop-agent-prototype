# novotech_functional_training.py
import re
import io
import os
import streamlit as st
import pandas as pd
from typing import List, Set

st.set_page_config(page_title="Novotech Functional Training", layout="wide")
st.title("Novotech Functional Training")

EXCEL_PATH = "/mnt/data/PR_Audience_LearningItem_Export.xlsx"

if not os.path.exists(EXCEL_PATH):
    st.error(f"Expected Excel file not found at {EXCEL_PATH}")
    st.stop()

# -------------------------
# Helpers to parse Column B
# -------------------------
# Patterns to capture the parenthesized lists after "Role : Any of : ( ... )" and
# "Organisation : Any of : ( ... )". We'll be tolerant of spaces and case.
ROLE_RE = re.compile(r"Role\s*:\s*Any of\s*:\s*\(([^)]*)\)", flags=re.I)
ORG_RE = re.compile(r"Organisation\s*:\s*Any of\s*:\s*\(([^)]*)\)", flags=re.I)

def split_items(s: str) -> List[str]:
    """Split comma-separated items inside parentheses; trim whitespace and empty items."""
    if not isinstance(s, str) or s.strip() == "":
        return []
    parts = [p.strip() for p in re.split(r",\s*", s) if p is not None]
    parts = [p for p in parts if p != ""]
    return parts

def extract_roles(cell: str) -> List[str]:
    """Return list of role names extracted from the Role(...) clause in Column B."""
    if not isinstance(cell, str):
        return []
    roles = []
    for m in ROLE_RE.finditer(cell):
        roles += split_items(m.group(1))
    # final cleanup: remove trailing/leading parentheses/spurious characters
    roles = [re.sub(r"^[\(\)\s]+|[\(\)\s]+$", "", r) for r in roles]
    return sorted(set([r for r in roles if r]))

def extract_org_groups_practices(cell: str):
    """
    From Organisation(...) return two lists: groups, practices.
    Items in Organisation(...) may have suffixes like '(Group)' or '(Practice)'.
    We'll detect those and strip the suffix.
    """
    if not isinstance(cell, str):
        return [], []
    groups = []
    practices = []
    for m in ORG_RE.finditer(cell):
        items = split_items(m.group(1))
        for it in items:
            it = it.strip()
            # detect trailing "(Group)" or "(Practice)" (case-insensitive)
            g_match = re.match(r"^(.*?)(?:\s*\(\s*Group\s*\)\s*)$", it, flags=re.I)
            p_match = re.match(r"^(.*?)(?:\s*\(\s*Practice\s*\)\s*)$", it, flags=re.I)
            if g_match:
                groups.append(g_match.group(1).strip())
            elif p_match:
                practices.append(p_match.group(1).strip())
            else:
                # if no explicit suffix, we can't be sure; ignore ambiguous entries
                # (alternatively, we could add to both or one — keep conservative)
                pass
    return sorted(set([g for g in groups if g])), sorted(set([p for p in practices if p]))

# -------------------------
# Load Excel and columns A..E
# -------------------------
# Read entire sheet (first sheet)
xls = pd.ExcelFile(EXCEL_PATH)
sheet_name = xls.sheet_names[0]
raw = pd.read_excel(EXCEL_PATH, sheet_name=sheet_name, header=None, dtype=object)

# Columns expected: A=0 Prescriptive rule, B=1 Member selection criteria, C=2 Course ID,
# D=3 Course Title, E=4 Curriculum Title
# Ensure at least 5 columns exist
max_needed_col = 5
if raw.shape[1] < max_needed_col:
    st.error(f"Expected at least {max_needed_col} columns (A-E) in the sheet. Found {raw.shape[1]}.")
    st.stop()

prescriptive = raw.iloc[:, 0].astype(object).fillna("").astype(str)
member_criteria = raw.iloc[:, 1].astype(object).fillna("").astype(str)
course_id_col = raw.iloc[:, 2].astype(object).fillna("").astype(str)
course_title_col = raw.iloc[:, 3].astype(object).fillna("").astype(str)
curriculum_title_col = raw.iloc[:, 4].astype(object).fillna("").astype(str)

nrows = len(raw)

# -------------------------
# Group rows by Prescriptive rule value
# -------------------------
rule_to_indices = {}
for idx, rule in enumerate(prescriptive):
    key = rule.strip()
    rule_to_indices.setdefault(key, []).append(idx)

# We'll build a list of output records (dicts) with fields:
# Title, ID, Type ('course'|'curriculum'), Groups (list), Practices (list), Roles (list)
records = []

# Helper to parse a single row's selection criteria into groups/practices/roles
def parse_row_attributes(index: int):
    cell = member_criteria[index]
    roles = extract_roles(cell)
    groups, practices = extract_org_groups_practices(cell)
    return groups, practices, roles

# Iterate through each distinct prescriptive rule key
for rule_key, indices in rule_to_indices.items():
    if rule_key != "" and len(indices) > 1:
        # Treat as curriculum + its course rows
        # Curriculum title: prefer first non-empty curriculum_title_col among indices
        curr_title = ""
        for i in indices:
            t = curriculum_title_col[i]
            if isinstance(t, str) and t.strip() != "":
                curr_title = t.strip()
                break
        # If still empty, fallback to rule_key as curriculum title (safe fallback)
        if not curr_title:
            curr_title = rule_key

        # Aggregate attributes from all rows in this block
        agg_groups: Set[str] = set()
        agg_practices: Set[str] = set()
        agg_roles: Set[str] = set()
        for i in indices:
            g, p, r = parse_row_attributes(i)
            agg_groups.update(g)
            agg_practices.update(p)
            agg_roles.update(r)

        # Add curriculum record (ID empty)
        records.append({
            "Title": curr_title,
            "ID": "",
            "Type": "curriculum",
            "Groups": sorted(list(agg_groups)),
            "Practices": sorted(list(agg_practices)),
            "Roles": sorted(list(agg_roles)),
            "SourceIndices": indices  # optional debug info
        })

        # Add individual course records for each row
        for i in indices:
            title = course_title_col[i].strip() if isinstance(course_title_col[i], str) else ""
            cid = course_id_col[i].strip() if isinstance(course_id_col[i], str) else ""
            g, p, r = parse_row_attributes(i)
            # If a course row lacks role/group/practice, it may still be filtered by the curriculum aggregates if needed.
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
        # Single occurrence -> treat as course
        i = indices[0]
        title = course_title_col[i].strip() if isinstance(course_title_col[i], str) else ""
        cid = course_id_col[i].strip() if isinstance(course_id_col[i], str) else ""
        g, p, r = parse_row_attributes(i)
        # If this single-row rule actually has a curriculum title in col E but rule is unique,
        # we still treat it as course per your instruction (Title from D).
        records.append({
            "Title": title if title else curriculum_title_col[i] if curriculum_title_col[i] else prescriptive[i],
            "ID": cid,
            "Type": "course",
            "Groups": sorted(list(g)),
            "Practices": sorted(list(p)),
            "Roles": sorted(list(r)),
            "SourceIndex": i
        })

# -------------------------
# Create DataFrame from records
# -------------------------
out_df = pd.DataFrame(records)

# Normalize list fields to empty lists where NaN
for col in ["Groups", "Practices", "Roles"]:
    if col not in out_df.columns:
        out_df[col] = [[] for _ in range(len(out_df))]
    else:
        out_df[col] = out_df[col].apply(lambda v: v if isinstance(v, list) else ([] if pd.isna(v) else [v]))

# -------------------------
# Build filter value lists (alphabetical)
# -------------------------
all_practices = sorted({p for row in out_df["Practices"] for p in row}, key=lambda x: x.lower())
all_groups = sorted({g for row in out_df["Groups"] for g in row}, key=lambda x: x.lower())
all_roles = sorted({r for row in out_df["Roles"] for r in row}, key=lambda x: x.lower())

# -------------------------
# UI: left filters (unchecked by default) + right table
# -------------------------
left_col, right_col = st.columns([1, 3])

# Utility to render checkbox list unchecked by default and return selected set
def checkbox_list_unchecked(label: str, options: List[str], key_prefix: str) -> List[str]:
    st.markdown(f"**{label}**")
    selected = []
    for i, opt in enumerate(options):
        key = f"{key_prefix}__{i}"
        # default unchecked
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
    st.write("Note: leave all checkboxes unchecked to include all values for that filter.")
    st.write("Filter logic: OR within a filter (multiple checked = include rows matching any), AND across filters.")

# -------------------------
# Apply filters to out_df
# -------------------------
def row_matches_filter_list(row_list: List[str], selected_list: List[str]) -> bool:
    """
    If selected_list is empty -> treat as match (no restriction).
    Otherwise, return True if any element in row_list is in selected_list.
    """
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
# Display table on the right
# -------------------------
with right_col:
    st.subheader("Filtered Learning Items")
    if filtered_df.empty:
        st.info("No items match the current filters.")
    else:
        # Show only Title, ID, Type as requested
        display_df = filtered_df[["Title", "ID", "Type"]].copy()
        # Clean empty ID to be truly empty instead of '': already set but normalize
        display_df["ID"] = display_df["ID"].replace("", "")
        st.dataframe(display_df.reset_index(drop=True), use_container_width=True)

        # CSV download
        csv_buf = io.StringIO()
        display_df.to_csv(csv_buf, index=False)
        st.download_button(
            label="Download visible items as CSV",
            data=csv_buf.getvalue().encode(),
            file_name="novotech_functional_training_filtered.csv",
            mime="text/csv"
        )

# -------------------------
# Optional: show counts and a small sample of attributes (for debugging)
# -------------------------
st.markdown("---")
st.write(f"Total parsed items: {len(out_df)} (courses + curricula). Showing {len(filtered_df)} after filters.")
if st.checkbox("Show debug: display first 50 parsed rows with attributes", value=False):
    debug_df = out_df.head(50).copy()
    # collapse lists to comma strings for readability
    for c in ["Groups", "Practices", "Roles"]:
        debug_df[c] = debug_df[c].apply(lambda L: ", ".join(L) if isinstance(L, list) else "")
    st.dataframe(debug_df.reset_index(drop=True), use_container_width=True)
