"""Sample project: Excel sales dashboard builder (openpyxl).

Reads a CSV of sales transactions (date, region, category, revenue) and
builds a management dashboard workbook with:

  * "Dashboard" sheet — KPI summary computed with real Excel formulas
    (total revenue, average monthly revenue, best category, best month),
    plus a native monthly-revenue column chart and a revenue-by-category
    pie chart.
  * "Summary" sheet — monthly and per-category totals via SUMIF formulas,
    feeding the KPIs and the charts.
  * "Data" sheet — the raw transactions as a formatted Excel Table,
    with a Month helper column (TEXT formula) used by the SUMIFs.

Demonstrates: reading CSV into a workbook, Excel Tables, SUMIF/INDEX/MATCH
formulas, native openpyxl BarChart/PieChart, number formats, cell styling.

Usage:
    pip install openpyxl
    python build_dashboard.py --input sales_data.csv --output dashboard.xlsx

Point it at your own CSV with the same four columns
(date, region, category, revenue) and it builds the same dashboard.
"""
import argparse
import csv
from datetime import date

from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, color="1F4E79", size=18)
KPI_LABEL_FONT = Font(bold=True, size=12, color="595959")
KPI_VALUE_FONT = Font(bold=True, size=14, color="1F4E79")
CURRENCY_FMT = '"$"#,##0'
THIN_BORDER = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)


def read_transactions(path):
    rows = []
    with open(path, newline="", encoding="utf-8") as f:
        for rec in csv.DictReader(f):
            rows.append((
                date.fromisoformat(rec["date"].strip()),
                rec["region"].strip(),
                rec["category"].strip(),
                float(rec["revenue"]),
            ))
    rows.sort(key=lambda r: r[0])
    return rows


def style_header(ws, row, ncols):
    for col in range(1, ncols + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER


def build_data_sheet(wb, rows):
    ws = wb.active
    ws.title = "Data"
    headers = ["Date", "Region", "Category", "Revenue", "Month"]
    ws.append(headers)
    style_header(ws, 1, len(headers))
    for i, (d, region, category, revenue) in enumerate(rows, start=2):
        ws.cell(row=i, column=1, value=d).number_format = "yyyy-mm-dd"
        ws.cell(row=i, column=2, value=region)
        ws.cell(row=i, column=3, value=category)
        ws.cell(row=i, column=4, value=revenue).number_format = CURRENCY_FMT
        ws.cell(row=i, column=5, value=f'=TEXT(A{i},"yyyy-mm")')
    last = len(rows) + 1
    tab = Table(displayName="SalesData", ref=f"A1:E{last}")
    tab.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium9", showRowStripes=True)
    ws.add_table(tab)
    for col, width in zip("ABCDE", (12, 12, 18, 12, 10)):
        ws.column_dimensions[col].width = width
    ws.sheet_view.showGridLines = False
    return ws, last


def build_summary_sheet(wb, last_row):
    ws = wb.create_sheet("Summary")
    # --- monthly totals ---
    ws["A1"] = "Month"
    ws["B1"] = "Revenue"
    style_header(ws, 1, 2)
    for m in range(1, 13):
        r = m + 1
        ws.cell(row=r, column=1, value=f"2025-{m:02d}")
        ws.cell(row=r, column=2,
                value=f"=SUMIF(Data!$E$2:$E${last_row},$A{r},Data!$D$2:$D${last_row})"
                ).number_format = CURRENCY_FMT
    # --- category totals ---
    ws["D1"] = "Category"
    ws["E1"] = "Revenue"
    style_header(ws, 1, 2)  # styles A1:B1 again, harmless
    for col in (4, 5):
        ws.cell(row=1, column=col).fill = HEADER_FILL
        ws.cell(row=1, column=col).font = HEADER_FONT
        ws.cell(row=1, column=col).alignment = Alignment(horizontal="center")
    for i, cat in enumerate(["Electronics", "Furniture",
                             "Office Supplies", "Accessories"], start=2):
        ws.cell(row=i, column=4, value=cat)
        ws.cell(row=i, column=5,
                value=f"=SUMIF(Data!$C$2:$C${last_row},$D{i},Data!$D$2:$D${last_row})"
                ).number_format = CURRENCY_FMT
    for col, width in zip("ABDE", (10, 14, 18, 14)):
        ws.column_dimensions[col].width = width
    ws.sheet_view.showGridLines = False
    return ws


def build_dashboard_sheet(wb, last_row):
    ws = wb.create_sheet("Dashboard", 0)  # first tab
    ws["A1"] = "Sales Dashboard — 2025"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = "Built from sales_data.csv with Python + openpyxl"
    ws["A2"].font = Font(italic=True, color="808080", size=10)

    # --- KPI summary (real Excel formulas) ---
    ws["A4"] = "Total Revenue"
    ws["B4"] = f"=SUM(Data!$D$2:$D${last_row})"
    ws["A5"] = "Average Monthly Revenue"
    ws["B5"] = "=B4/12"
    ws["A6"] = "Best Category"
    ws["B6"] = ("=INDEX(Summary!$D$2:$D$5,"
                "MATCH(MAX(Summary!$E$2:$E$5),Summary!$E$2:$E$5,0))")
    ws["A7"] = "Best Month"
    ws["B7"] = ("=INDEX(Summary!$A$2:$A$13,"
                "MATCH(MAX(Summary!$B$2:$B$13),Summary!$B$2:$B$13,0))")
    for r in range(4, 8):
        ws.cell(row=r, column=1).font = KPI_LABEL_FONT
        ws.cell(row=r, column=1).border = THIN_BORDER
        ws.cell(row=r, column=2).font = KPI_VALUE_FONT
        ws.cell(row=r, column=2).border = THIN_BORDER
        ws.cell(row=r, column=2).alignment = Alignment(horizontal="center")
    ws["B4"].number_format = CURRENCY_FMT
    ws["B5"].number_format = CURRENCY_FMT
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 22

    # --- monthly revenue column chart ---
    bar = BarChart()
    bar.type = "col"
    bar.title = "Monthly Revenue"
    bar.y_axis.title = "Revenue"
    bar.y_axis.numFmt = CURRENCY_FMT
    bar.style = 10
    bar.height = 7.5
    bar.width = 14
    data = Reference(wb["Summary"], min_col=2, min_row=1, max_row=13)
    cats = Reference(wb["Summary"], min_col=1, min_row=2, max_row=13)
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)
    bar.dataLabels = DataLabelList()
    bar.dataLabels.showVal = False
    bar.shape = 4
    ws.add_chart(bar, "A9")

    # --- revenue by category pie chart ---
    pie = PieChart()
    pie.title = "Revenue by Category"
    pie.style = 10
    pie.height = 7.5
    pie.width = 12
    pdata = Reference(wb["Summary"], min_col=5, min_row=1, max_row=5)
    pcats = Reference(wb["Summary"], min_col=4, min_row=2, max_row=5)
    pie.add_data(pdata, titles_from_data=True)
    pie.set_categories(pcats)
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True
    ws.add_chart(pie, "I9")

    ws.sheet_view.showGridLines = False
    return ws


def main():
    ap = argparse.ArgumentParser(description="Build an Excel sales dashboard.")
    ap.add_argument("--input", default="sales_data.csv")
    ap.add_argument("--output", default="dashboard.xlsx")
    args = ap.parse_args()

    rows = read_transactions(args.input)
    wb = Workbook()
    _, last_row = build_data_sheet(wb, rows)
    build_summary_sheet(wb, last_row)
    build_dashboard_sheet(wb, last_row)
    wb.save(args.output)
    print(f"Dashboard written to {args.output} "
          f"({len(rows)} transactions, {wb.sheetnames})")


if __name__ == "__main__":
    main()
