# Invoice → Excel (Hotel Shanth / Anuradha Kagod)

Upload the scanned KSBCL invoice PDFs → get one workbook with **Summary**, **Purchase <Month>** and **Sales <Month>**,
formatted exactly like `template.xlsx` (your August workbook).

## One-time setup
1. Python 3.10+
2. Install the two OCR programs and make sure they are on PATH:
   * Tesseract OCR  (Windows: UB-Mannheim installer; Mac: `brew install tesseract`; Linux: `apt install tesseract-ocr`)
   * Poppler        (Windows: poppler-windows zip; Mac: `brew install poppler`; Linux: `apt install poppler-utils`)
3. `pip install -r requirements.txt`

## Use
* Web screen:  `python app.py`  → open http://127.0.0.1:5000 → upload PDFs → check the review table → *Create Excel workbook*
* Command line: `python invoice2xlsx.py inv1.pdf inv2.pdf inv3.pdf inv4.pdf -o Sep.xlsx --sales-start 919`

## What it does
* **Purchase sheet** – one block per invoice (4 invoices = 4 blocks), items/CB/Btls/rate/amount, SUM total row.
* **Sales sheet** – each purchase invoice is split into two sales invoices (first half of rows = ceil(n/2)),
  numbered from `--sales-start` (default **911**), dated 2 days after the first purchase then every 3 days.
  Amount = purchase amount × 1.10 (live formulas pointing at the Purchase sheet, so editing a purchase row updates sales).
* **Summary sheet** – purchase totals and sales totals, same layout as August.

## Accuracy checks (shown on the review screen)
* every row: amount = CB × rate + Btls × rate ÷ bottles-per-case
* sum of rows = invoice total printed at the bottom of the invoice
* CB / Btls totals = permit page totals
Rows that fail are highlighted yellow/red. Fix them on screen before creating the workbook.
Items it has not seen before are flagged "new item"; once you create a workbook they are remembered (`learned_catalog.json`),
so the same names/rates are recognised next month.

## Files
`engine.py` (OCR + parsing + workbook builder) · `app.py` (web screen) · `invoice2xlsx.py` (command line) · `template.xlsx` (formatting + known items)
