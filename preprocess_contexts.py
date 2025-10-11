import pandas as pd
import os

# -------------------------------
# CONFIG
# -------------------------------
DATA_PATH = "data/Novotech_SOP_Matrix.xlsx"
OUTPUT_DIR = "output"

# Create output folder if it doesn't exist
os.makedirs(OUTPUT_DIR, exist_ok=True)

# -------------------------------
# LOAD EXCEL
# -------------------------------
# We use header=None because first 3 rows are merged/complex
df = pd.read_excel(DATA_PATH, sheet_name=0, header=None)

# -------------------------------
# EXTRACT GROUPS AND ROLES
# -------------------------------
# Row 0 (first row) columns E onward = Group names (merged)
groups = df.iloc[0, 4:]

# Row 2 (third row) columns E onward = Roles
roles = df.iloc[2, 4:]

role_info = []
for col, role in enumerate(roles, start=4):  # Column E = index 4
    group = groups[col]
    role_info.append({"role": role, "group": group, "col": col})

# -------------------------------
# EXTRACT SOPs PER ROLE
# -------------------------------
sops_per_role = {}

for info in role_info:
    role_name = info["role"]
    col = info["col"]
    sops = []
    
    for idx, row in df.iloc[3:].iterrows():  # SOPs start from row 4
        assigned = row[col]
        if str(assigned).strip() in ["1", "2", "3"]:  # Only assigned SOPs
            sops.append({
                "Business Unit": row[0],
                "SOP Type": row[1],
                "Number": row[2],
                "Title": row[3],
                "Group": info["group"]
            })
    
    sops_per_role[role_name] = sops

# -------------------------------
# WRITE TXT FILES
# -------------------------------
import re

def sanitize_filename(name):
    return re.sub(r'[<>:"/\\|?*]', '_', str(name))

for role, sops in sops_per_role.items():
    safe_role_name = sanitize_filename(role)
    file_path = os.path.join(OUTPUT_DIR, f"{safe_role_name}.txt")
    with open(file_path, "w", encoding="utf-8") as f:
        f.write(f"SOPs for {role}:\n\n")
        for sop in sops:
            f.write(
                f"- Business Unit: {sop['Business Unit']} | "
                f"SOP Type: {sop['SOP Type']} | "
                f"Number: {sop['Number']} | "
                f"Title: {sop['Title']} | "
                f"Group: {sop['Group']}\n"
            )