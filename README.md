# pdffiller

A small Python tool that batch-fills a fillable PDF form from rows of a spreadsheet.

It reads field values out of an Excel workbook and writes one filled PDF per data
row, so a single template form can be produced in bulk.

## How it works

- The PDF template is `test.pdf`; its form fields are discovered with `fillpdf`.
- The workbook is `book1.xlsx`, read with `openpyxl`:
  - row 2 holds the PDF field names (the column headers that map a column to a form field),
  - rows 3+ hold one record per row.
- For each form field, a single-field sample PDF is written to `testfields/<field>.pdf`
  (useful for identifying which visual field a name maps to).
- For each data row, a filled PDF is written to `output/OutputFilename<col1>.pdf`.

## Requirements

- Python 3
- `openpyxl`
- `fillpdf`

```sh
pip install openpyxl fillpdf
```

## Run

Place `test.pdf` and `book1.xlsx` next to `main.py`, then:

```sh
mkdir -p output testfields
python3 main.py
```
