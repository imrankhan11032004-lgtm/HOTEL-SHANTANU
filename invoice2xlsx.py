"""Command line:  python invoice2xlsx.py invoice1.pdf invoice2.pdf ... -o Aug.xlsx [--sales-start 911] [--markup 10]"""
import argparse, sys
import engine

ap = argparse.ArgumentParser()
ap.add_argument("pdfs", nargs="+")
ap.add_argument("-o", "--out", default="Liquor_Purchase_Sales.xlsx")
ap.add_argument("--sales-start", type=int, default=911, help="first sales invoice number (default 911)")
ap.add_argument("--markup", type=float, default=10, help="sales mark-up in percent (default 10)")
ap.add_argument("--first-sale-offset", type=int, default=2, help="days after the first purchase for the first sale")
ap.add_argument("--gap", type=int, default=3, help="days between sales invoices")
a = ap.parse_args()

cat = engine.load_catalog()
invs, problems = [], 0
for p in a.pdfs:
    print("Reading", p, "...", flush=True)
    inv = engine.read_invoice(p, catalog=cat)
    invs.append(inv)
    print(f"  {inv['invoice_no']}  {inv['date']}  {len(inv['items'])} items  total {inv['total_items']:,.2f}")
    for w in inv["warnings"]: print("  WARNING:", w); problems += 1
    for it in inv["items"]:
        if it["flags"]: print(f"  check row {it['sl']}: {it['name'][:50]} -> {'; '.join(it['flags'])}")
info = engine.build_workbook(invs, a.out, sales_start=a.sales_start, markup=a.markup / 100,
                             first_sale_offset=a.first_sale_offset, sale_gap_days=a.gap)
engine.learn([it for inv in invs for it in inv["items"] if it["rate"] and not it["flags"]])
print(f"\nSaved {a.out}  (sales invoices {info['sales_numbers'][0]}-{info['sales_numbers'][-1]})")
if problems: print("Some invoices did not reconcile - open the file and compare the flagged rows with the paper invoice.")
