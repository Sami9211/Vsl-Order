"""Parse a Vape Supplier Ltd (VSL) invoice PDF into a verified CSV of line items.

Usage: python scripts/parse_invoice.py <invoice.pdf> [<invoice.pdf> ...]

For every invoice it checks that quantity x unit price equals each line total and
that the line totals add up to the invoice subtotal, and refuses to write the CSV
if anything does not reconcile.
"""
import csv
import re
import subprocess
import sys
from decimal import Decimal
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT_DIR = ROOT / "data" / "invoices"

ITEM_ROW = re.compile(r"^(.*?)\s{2,}(\d+)\s+£([\d,]+\.\d\d)\s+(\d+)%\s+£([\d,]+\.\d\d)$")
FOOTER = "Vape Supplier Ltd ‧ Phone"
ORDER_NO = re.compile(r"Order Number:\s+(\d+)")
ISSUE_DATE = re.compile(r"Issue Date:\s+(.+?)\s*$")
SUBTOTAL = re.compile(r"^\s*Subtotal\s+£([\d,]+\.\d\d)\s*$")
DESC = re.compile(r"^(?:NEW!\s+)?(Hayati .+?)(?:\s+20mg)?\s*\[(.+)\s/\s(\d+mg)\]\s*-$")

PRODUCT_LINES = {
    "Hayati Pro Max+ 6K Prefilled Pod Kit": "Pro Max+ 6K Pod Kit",
    "Hayati Pro Max+ 6K Prefilled Replacement Pods": "Pro Max+ 6K Replacement Pods",
    "Hayati Pro Ultra+ 25K Prefilled Pod Kit": "Pro Ultra+ 25K Pod Kit",
    "Hayati Pro Ultra+ 25K Prefilled Pods": "Pro Ultra+ 25K Replacement Pods",
}


def money(s):
    return Decimal(s.replace(",", ""))


def parse(pdf):
    text = subprocess.run(["pdftotext", "-layout", str(pdf), "-"],
                          capture_output=True, text=True, check=True).stdout
    lines = text.splitlines()
    order_no = issue_date = client = subtotal = None
    items, current = [], None
    for i, raw in enumerate(lines):
        line = raw.strip()
        if m := ORDER_NO.search(raw):
            order_no = m.group(1)
            issue_date = ISSUE_DATE.search(raw).group(1)
        if raw.lstrip().startswith("Supplier:") and "Client:" in raw:
            client = lines[i + 1].split("  ")[-1].strip()
        if m := SUBTOTAL.match(raw):
            subtotal = money(m.group(1))
        if not line or FOOTER in line:
            continue
        if m := ITEM_ROW.match(line):
            if current is not None:
                raise ValueError(f"line {i+1}: previous description not finished: {current['description']!r}")
            current = {"description": m.group(1).strip(), "quantity": int(m.group(2)),
                       "price": money(m.group(3)), "total": money(m.group(5))}
        elif current is not None:
            current["description"] += " " + line
        else:
            continue
        if current["description"].endswith("] -"):
            items.append(current)
            current = None
    if current is not None:
        raise ValueError(f"unfinished item: {current}")

    errors = []
    for it in items:
        if it["quantity"] * it["price"] != it["total"]:
            errors.append(f"qty x price != total: {it}")
        m = DESC.match(it["description"])
        if not m or m.group(1) not in PRODUCT_LINES:
            errors.append(f"unrecognised description: {it['description']!r}")
            continue
        it["product_line"] = PRODUCT_LINES[m.group(1)]
        it["flavour"] = m.group(2).strip()
        it["strength"] = m.group(3)
    line_sum = sum(it["total"] for it in items)
    if subtotal is None or line_sum != subtotal:
        errors.append(f"line totals {line_sum} != invoice subtotal {subtotal}")
    if not order_no or not client:
        errors.append("missing order number / client")
    keys = [(it.get("product_line"), it.get("flavour"), it.get("strength")) for it in items]
    if len(keys) != len(set(keys)):
        errors.append("duplicate product on the same invoice")
    if errors:
        raise ValueError(f"{pdf}:\n  " + "\n  ".join(errors))
    return order_no, issue_date, client, subtotal, items


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    for pdf in sys.argv[1:]:
        order_no, issue_date, client, subtotal, items = parse(pdf)
        out = OUT_DIR / f"{order_no}.csv"
        with out.open("w", newline="") as f:
            w = csv.writer(f)
            w.writerow(["invoice", "issue_date", "client", "product_line", "flavour",
                        "strength", "quantity", "invoice_description"])
            for it in items:
                w.writerow([order_no, issue_date, client, it["product_line"], it["flavour"],
                            it["strength"], it["quantity"], it["description"]])
        units = sum(it["quantity"] for it in items)
        print(f"{order_no} ({issue_date}, {client}): {len(items)} lines, {units} units, "
              f"subtotal £{subtotal} reconciled -> {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
