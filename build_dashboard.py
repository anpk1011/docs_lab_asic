import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import PieChart, BarChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows

INPUT = "coverity_report.xlsx"
OUTPUT = "coverity_dashboard.xlsx"

# Load summary
df_summary = pd.read_excel(INPUT, sheet_name="Summary")

# Copy workbook
wb = load_workbook(INPUT)

# Remove old dashboard sheets
for s in wb.sheetnames:
    if s.startswith("Dashboard"):
        del wb[s]

# --- CREATE MAIN DASHBOARD ---
ws = wb.create_sheet("Dashboard_Main")

# Aggregate first snapshot
df_old = df_summary[df_summary["Snapshot"]=="first"].groupby("Category")["Count"].sum().reset_index()

# Aggregate last snapshot
df_last = df_summary[df_summary["Snapshot"]=="last"].groupby("Category")["Count"].sum().reset_index()

# Write old snapshot table
ws.append(["Category (OLD)", "Count"])
for r in df_old.values:
    ws.append(list(r))

# Add pie chart for OLD
pie1 = PieChart()
pie1.title = "Total Categories - Snapshot FIRST"
labels = Reference(ws, min_col=1, min_row=2, max_row=len(df_old)+1)
data = Reference(ws, min_col=2, min_row=1, max_row=len(df_old)+1)
pie1.add_data(data, titles_from_data=True)
pie1.set_categories(labels)
ws.add_chart(pie1, "D2")


# Write last snapshot table
start = len(df_old) + 5
ws.cell(row=start-1, column=1, value="Category (LAST)")
ws.cell(row=start-1, column=2, value="Count")

for i, r in enumerate(df_last.values, start=start):
    ws.append(list(r))

# Add pie chart for LAST
pie2 = PieChart()
pie2.title = "Total Categories - Snapshot LAST"
labels2 = Reference(ws, min_col=1, min_row=start+1, max_row=start+len(df_last))
data2   = Reference(ws, min_col=2, min_row=start,   max_row=start+len(df_last))
pie2.add_data(data2, titles_from_data=True)
pie2.set_categories(labels2)
ws.add_chart(pie2, "D20")

# --- CREATE PROJECT DETAIL DASHBOARDS ---
projects = df_summary["Project"].unique()

for proj in projects:
    ws2 = wb.create_sheet(f"Dashboard_{proj}")

    df_p = df_summary[df_summary["Project"] == proj]

    # Write raw data for project
    for r in dataframe_to_rows(df_p, index=False, header=True):
        ws2.append(r)

    # BAR CHART compare first vs last
    chart = BarChart()
    chart.title = f"{proj} - Snapshot Comparison"
    chart.x_axis.title = "Category"
    chart.y_axis.title = "Count"

    # Data range
    data_ref = Reference(ws2, min_col=4, min_row=1, max_col=4, max_row=len(df_p)+1)
    cat_ref = Reference(ws2, min_col=3, min_row=2, max_row=len(df_p)+1)

    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cat_ref)
    ws2.add_chart(chart, "H2")

wb.save(OUTPUT)

print("DONE → coverity_dashboard.xlsx generated.")
