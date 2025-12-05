import json
import pandas as pd

INPUT_FILE = "defects.json"
OUTPUT_FILE = "coverity_report.xlsx"

raw_records = []

# read JSON in defects.json
with open(INPUT_FILE, "r") as f:
    for line in f:
        if not line.strip():
            continue
        block = json.loads(line)

        project = block["project"]
        proj_key = block["project_key"]
        view_id = int(block["view_id"])
        snapshot = "first" if view_id == 10099 else "last"

        for row in block["rows"]:
            rec = {}
            rec.update(row)
            rec["Project"] = project
            rec["ProjectKey"] = proj_key
            rec["Snapshot"] = snapshot

            # Normalize Category
            cat = rec.get("Category", "")
            if isinstance(cat, list):
                cat = " ".join(cat)
            cat = " ".join(str(cat).split())
            rec["Category"] = cat

            # Convert Count → int
            rec["Count"] = int(rec.get("Count", 0) or 0)

            raw_records.append(rec)

# Convert raw table → DataFrame
df_raw = pd.DataFrame(raw_records)

# Fill missing columns if some rows lack fields
for col in ["CID", "Classification", "Component", "Action", "Category", "Count"]:
    if col not in df_raw.columns:
        df_raw[col] = None

# Build summary pivot
df_summary = df_raw.groupby(["Project", "Snapshot", "Category"])["Count"].sum().reset_index()

# Xuất excel
with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    df_raw.to_excel(writer, sheet_name="Defects_Raw", index=False)
    df_summary.to_excel(writer, sheet_name="Summary", index=False)

print("===================================================")
print(f" DONE! File Excel created: {OUTPUT_FILE} ")
print("===================================================")
