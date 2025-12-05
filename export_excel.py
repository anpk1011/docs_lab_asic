import json
import pandas as pd

INPUT_FILE = "defects.json"
OUTPUT_FILE = "coverity_report.xlsx"

records = []

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

            # MAP lại key đúng chuẩn Coverity
            rec["CID"] = row.get("cid")
            rec["Classification"] = row.get("classification")
            rec["Component"] = row.get("displayComponent")
            rec["Action"] = row.get("action")

            # Category trong API
            cat = row.get("displayCategory", "")
            if isinstance(cat, list):
                cat = " ".join(cat)
            rec["Category"] = " ".join(str(cat).split())

            # Count
            rec["Count"] = int(row.get("occurrenceCount", 0))

            rec["Project"] = project
            rec["ProjectKey"] = proj_key
            rec["Snapshot"] = snapshot

            records.append(rec)

df_raw = pd.DataFrame(records)

# Summary for charts
df_summary = df_raw.groupby(["Project", "Snapshot", "Category"])["Count"].sum().reset_index()

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    df_raw.to_excel(writer, sheet_name="Defects_Raw", index=False)
    df_summary.to_excel(writer, sheet_name="Summary", index=False)

print("DONE → coverity_report.xlsx")
