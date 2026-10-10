"""Build VSL_Order_Totals.xlsx from every verified invoice CSV in data/invoices/.

Usage: python scripts/build_workbook.py
"""
import csv
from collections import defaultdict
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
IN_DIR = ROOT / "data" / "invoices"
OUT = ROOT / "VSL_Order_Totals.xlsx"

LINE_ORDER = ["Pro Max+ 6K Pod Kit", "Pro Max+ 6K Replacement Pods",
              "Pro Ultra+ 25K Pod Kit", "Pro Ultra+ 25K Replacement Pods"]

FONT = "Arial"
BOLD = Font(name=FONT, bold=True)
NORMAL = Font(name=FONT)
HEADER_FONT = Font(name=FONT, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
TOTAL_FILL = PatternFill("solid", fgColor="DDEBF7")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
QTY_FMT = '#,##0;-#,##0;"-"'


def load_invoices():
    invoices = []
    for path in IN_DIR.glob("*.csv"):
        with path.open(newline="") as f:
            rows = list(csv.DictReader(f))
        first = rows[0]
        invoices.append({
            "number": first["invoice"],
            "date": datetime.strptime(first["issue_date"].replace(".", ""), "%b %d, %Y").date(),
            "client": first["client"],
            "rows": rows,
        })
    invoices.sort(key=lambda inv: (inv["date"], inv["number"]))
    return invoices


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font, cell.fill, cell.border = HEADER_FONT, HEADER_FILL, BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def main():
    invoices = load_invoices()
    qty = defaultdict(lambda: defaultdict(int))  # (line, flavour, strength) -> invoice -> qty
    for inv in invoices:
        for r in inv["rows"]:
            qty[(r["product_line"], r["flavour"], r["strength"])][inv["number"]] += int(r["quantity"])
    products = sorted(qty, key=lambda k: (LINE_ORDER.index(k[0]), k[1].lower(), k[2]))

    wb = Workbook()

    # --- Order Totals -------------------------------------------------------
    ws = wb.active
    ws.title = "Order Totals"
    n_inv = len(invoices)
    inv_cols = list(range(4, 4 + n_inv))
    total_col = 4 + n_inv
    tl = get_column_letter(total_col)
    headers = ["Product Line", "Flavour", "Strength"] + \
              [f"Invoice {inv['number']}\n{inv['date']:%d %b %Y}" for inv in invoices] + ["TOTAL QTY"]
    ws.append(headers)
    style_header(ws, 1, len(headers))
    ws.row_dimensions[1].height = 32
    first = 2
    for i, key in enumerate(products):
        r = first + i
        ws.cell(row=r, column=1, value=key[0])
        ws.cell(row=r, column=2, value=key[1])
        ws.cell(row=r, column=3, value=key[2])
        for c, inv in zip(inv_cols, invoices):
            v = qty[key].get(inv["number"])
            ws.cell(row=r, column=c, value=v)
        ws.cell(row=r, column=total_col,
                value=f"=SUM({get_column_letter(inv_cols[0])}{r}:{get_column_letter(inv_cols[-1])}{r})")
        for c in range(1, total_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.border, cell.font = BORDER, NORMAL
            if c >= 4:
                cell.number_format = QTY_FMT
        ws.cell(row=r, column=total_col).font = BOLD
    last = first + len(products) - 1

    # Summary by product line
    r = last + 2
    ws.cell(row=r, column=1, value="SUMMARY BY PRODUCT LINE").font = BOLD
    r += 1
    sum_first = r
    for line in LINE_ORDER:
        if not any(k[0] == line for k in products):
            continue
        ws.cell(row=r, column=1, value=line)
        ws.cell(row=r, column=2, value=f'=COUNTIF($A${first}:$A${last},A{r})&" flavours"')
        for c in inv_cols + [total_col]:
            col = get_column_letter(c)
            ws.cell(row=r, column=c, value=f"=SUMIF($A${first}:$A${last},$A{r},{col}${first}:{col}${last})")
        for c in range(1, total_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.border, cell.font, cell.fill = BORDER, BOLD, TOTAL_FILL
            if c >= 4:
                cell.number_format = QTY_FMT
        r += 1
    sum_last = r - 1
    ws.cell(row=r, column=1, value="GRAND TOTAL")
    for c in inv_cols + [total_col]:
        col = get_column_letter(c)
        ws.cell(row=r, column=c, value=f"=SUM({col}{first}:{col}{last})")
    for c in range(1, total_col + 1):
        cell = ws.cell(row=r, column=c)
        cell.border, cell.fill = BORDER, HEADER_FILL
        cell.font = HEADER_FONT
        if c >= 4:
            cell.number_format = QTY_FMT
    grand_row = r
    r += 1
    ws.cell(row=r, column=1, value="Check: product-line summary = grand total").font = Font(name=FONT, italic=True)
    ws.cell(row=r, column=total_col,
            value=f'=IF(SUM({tl}{sum_first}:{tl}{sum_last})={tl}{grand_row},"OK","MISMATCH")').font = BOLD

    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 48
    ws.column_dimensions["C"].width = 10
    for c in inv_cols + [total_col]:
        ws.column_dimensions[get_column_letter(c)].width = 15
    ws.freeze_panes = "D2"
    ws.auto_filter.ref = f"A1:{tl}{last}"

    # --- Invoices -------------------------------------------------------------
    wi = wb.create_sheet("Invoices")
    wi.append(["Invoice No", "Issue Date", "Billed To", "Line Items", "Units on Invoice",
               "Units in Order Totals", "Check"])
    style_header(wi, 1, 7)
    for i, (inv, c) in enumerate(zip(invoices, inv_cols)):
        r = i + 2
        col = get_column_letter(c)
        wi.cell(row=r, column=1, value=inv["number"])
        wi.cell(row=r, column=2, value=inv["date"]).number_format = "dd mmm yyyy"
        wi.cell(row=r, column=3, value=inv["client"])
        wi.cell(row=r, column=4, value=len(inv["rows"]))
        wi.cell(row=r, column=5, value=sum(int(x["quantity"]) for x in inv["rows"]))
        wi.cell(row=r, column=6, value=f"='Order Totals'!{col}{grand_row}")
        wi.cell(row=r, column=7, value=f'=IF(E{r}=F{r},"OK","MISMATCH")')
        for cc in range(1, 8):
            cell = wi.cell(row=r, column=cc)
            cell.border, cell.font = BORDER, NORMAL
            if cc in (4, 5, 6):
                cell.number_format = "#,##0"
        wi.cell(row=r, column=6).font = Font(name=FONT, color="008000")
    r = len(invoices) + 2
    wi.cell(row=r, column=1, value="TOTAL")
    for cc in (4, 5, 6):
        col = get_column_letter(cc)
        wi.cell(row=r, column=cc, value=f"=SUM({col}2:{col}{r-1})").number_format = "#,##0"
    wi.cell(row=r, column=7, value=f'=IF(E{r}=F{r},"OK","MISMATCH")')
    for cc in range(1, 8):
        cell = wi.cell(row=r, column=cc)
        cell.border, cell.font, cell.fill = BORDER, BOLD, TOTAL_FILL
    wi.cell(row=r + 2, column=1, value=(
        "Units on Invoice = sum of the quantity column printed on the PDF (each invoice was "
        "also reconciled: qty x price = line total, and line totals = invoice subtotal)."
    )).font = Font(name=FONT, italic=True)
    for col, w in zip("ABCDEFG", (14, 14, 22, 12, 16, 20, 12)):
        wi.column_dimensions[col].width = w

    # --- Invoice Lines (audit trail) -----------------------------------------
    wl = wb.create_sheet("Invoice Lines")
    wl.append(["Invoice No", "Product Line", "Flavour", "Strength", "Quantity",
               "Description exactly as on invoice"])
    style_header(wl, 1, 6)
    for inv in invoices:
        for x in inv["rows"]:
            wl.append([inv["number"], x["product_line"], x["flavour"], x["strength"],
                       int(x["quantity"]), x["invoice_description"]])
    for row in wl.iter_rows(min_row=2):
        for cell in row:
            cell.font, cell.border = NORMAL, BORDER
    for col, w in zip("ABCDEF", (12, 30, 46, 10, 10, 95)):
        wl.column_dimensions[col].width = w
    wl.freeze_panes = "A2"
    wl.auto_filter.ref = f"A1:F{wl.max_row}"

    wb.save(OUT)
    print(f"Wrote {OUT.name}: {len(invoices)} invoices, {len(products)} products")


if __name__ == "__main__":
    main()
