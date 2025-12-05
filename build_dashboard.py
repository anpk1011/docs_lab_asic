import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import PieChart, BarChart, Reference
from openpyxl.chart.label import DataLabelList

INPUT = "coverity_report.xlsx"
OUTPUT = "coverity_dashboard_optA.xlsx"

# -----------------------
# Force font = Arial
# -----------------------
def force_arial(chart):
    """
    Patch chart title and label fonts to Arial to avoid broken text
    when opening on Linux/LibreOffice/WPS.
    """
    try:
        # Fix chart title font
        if chart.title and chart.title.tx and chart.title.tx.rich:
            rPr = chart.title.tx.rich.p[0].r[0].rPr
            rPr.latin.typeface = "Arial"
            rPr.cs.typeface = "Arial"
            rPr.ea.typeface = "Arial"
    except:
        pass

# ---------------------------------
# Load data
# ---------------------------------
df_summary = pd.read_excel(INPUT, sheet_name="Summary")

# Beautify categories (main dashboard only)
def beautify(cat: str):
    if not isinstance(cat, str):
        return cat
    return " ".join(w.capitalize() for w in str(cat).split())

df_summary["Category"] = df_summary["Category"].apply(beautify)

# Copy workbook
wb = load_workbook(INPUT)

# Remove old dashboards
for s in list(wb.sheetnames):
    if s.startswith("Dashboard"):
        del wb[s]

# ---------------------------------
# MAIN DASHBOARD
# ---------------------------------
ws = wb.create_sheet("Dashboard_Main")

df_old = df_summary[df_summary["Snapshot"] == "first"].groupby("Category")["Count"].sum().reset_index()
df_last = df_summary[df_summary["Snapshot"] == "last"].groupby("Category")["Count"].sum().reset_index()

# Old table
ws.append(["Category (OLD)", "Count"])
for row in df_old.itertuples(index=False):
    ws.append(list(row))

# PIE FIRST
pie1 = PieChart()
pie1.title = "Total Categories - Snapshot FIRST"

labels = Reference(ws, min_col=1, min_row=2, max_row=1 + len(df_old))
data = Reference(ws, min_col=2, min_row=1, max_row=1 + len(df_old))

pie1.add_data(data, titles_from_data=True)
pie1.set_categories(labels)

force_arial(pie1)
ws.add_chart(pie1, "E2")

# LAST table
start = len(df_old) + 5
ws.cell(row=start - 1, column=1, value="Category (LAST)")
ws.cell(row=start - 1, column=2, value="Count")

for row in df_last.itertuples(index=False):
    ws.append(list(row))

# PIE LAST
pie2 = PieChart()
pie2.title = "Total Categories - Snapshot LAST"

labels2 = Reference(ws, min_col=1, min_row=start + 1, max_row=start + len(df_last))
data2 = Reference(ws, min_col=2, min_row=start, max_row=start + len(df_last))

pie2.add_data(data2, titles_from_data=True)
pie2.set_categories(labels2)

force_arial(pie2)
ws.add_chart(pie2, "E20")

# ---------------------------------
# DETAIL DASHBOARDS (TOTAL FIRST vs LAST ONLY)
# ---------------------------------
projects = df_summary["Project"].unique()

for proj in projects:
    ws2 = wb.create_sheet(f"Dashboard_{proj}")

    df_p = df_summary[df_summary["Project"] == proj]

    # Total defects
    total_first = int(df_p[df_p["Snapshot"] == "first"]["Count"].sum())
    total_last = int(df_p[df_p["Snapshot"] == "last"]["Count"].sum())
    delta_total = total_last - total_first

    # KPI
    ws2["A1"] = "Project"
    ws2["B1"] = proj

    ws2["A3"] = "Total FIRST"
    ws2["B3"] = total_first

    ws2["A4"] = "Total LAST"
    ws2["B4"] = total_last

    ws2["A5"] = "Delta (LAST - FIRST)"
    ws2["B5"] = delta_total

    # Chart table
    ws2.append([])
    ws2.append(["Snapshot", "Count"])
    ws2.append(["first", total_first])
    ws2.append(["last", total_last])

    # BAR chart
    bar = BarChart()
    bar.type = "col"
    bar.style = 10
    bar.title = f"{proj} – Total Defects Comparison"
    bar.y_axis.title = "Total Defects"
    bar.x_axis.title = "Snapshot"

    data_ref = Reference(ws2, min_col=2, min_row=8, max_row=9)
    cats_ref = Reference(ws2, min_col=1, min_row=8, max_row=9)

    bar.add_data(data_ref, titles_from_data=False)
    bar.set_categories(cats_ref)

    # Show labels on bars
    bar.dataLabels = DataLabelList()
    bar.dataLabels.showVal = True

    force_arial(bar)
    ws2.add_chart(bar, "E2")

# ---------------------------------
# SAVE OUTPUT
# ---------------------------------
wb.save(OUTPUT)
print("DONE →", OUTPUT)
