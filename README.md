# VSL Order Tracker

Running total of how many of each product has been ordered from Vape Supplier Ltd (VSL)
across all invoices for this order. Quantities only — no prices.

**Open [`VSL_Order_Totals.xlsx`](VSL_Order_Totals.xlsx).**

| Sheet | What it shows |
|---|---|
| Order Totals | One row per product (product line + flavour + strength), one column per invoice, and a **TOTAL QTY** column. Summary per product line and grand total at the bottom. |
| Invoices | Each invoice added so far, with its unit count and an OK/MISMATCH check against the totals sheet. |
| Invoice Lines | Every invoice line exactly as printed, for auditing. |

## Invoices included

| Invoice | Date | Line items | Units |
|---|---|---|---|
| 11156445 | 06 Oct 2026 | 181 | 4,445 |
| 11159830 | 09 Oct 2026 | 53 | 1,060 |
| **Total** | | **234** | **5,505** |

## Matching rules

Lines are matched across invoices by product line + flavour + strength. The `NEW!` prefix and the
extra `20mg` that sometimes appears before the flavour bracket are ignored, since they name the
same product. Flavour names are kept exactly as VSL prints them.

## Adding an invoice

```
python scripts/parse_invoice.py path/to/Invoice-XXXXXXXX.pdf   # writes data/invoices/XXXXXXXX.csv
python scripts/build_workbook.py                                # rebuilds VSL_Order_Totals.xlsx
```

The parser refuses an invoice unless every line's quantity × price equals its line total and the
line totals add up to the invoice subtotal, so a misread quantity can't slip through.
