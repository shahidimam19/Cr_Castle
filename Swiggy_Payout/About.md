# Swiggy Invoice HSN Extractor

A Python utility that scans all Swiggy PDF invoices in a folder and generates an Excel summary containing:

* Invoice Number
* PDF File Name
* Total Base Amount for HSN codes ending with **11**
* Total Base Amount for HSN codes not ending with **11**
* Calculated Subtotal

## Features

* Processes multiple PDF invoices in one run.
* Automatically extracts invoice numbers.
* Separates base amounts based on HSN code classification.
* Exports results to Excel (`.xlsx`).
* Works on Windows, Linux, and macOS.

---

## Example

### Input PDF

Invoice contains line items such as:

| Description     | HSN    | Base Amount |
| --------------- | ------ | ----------: |
| Service Fee     | 996211 |   28,883.43 |
| Collection Fee  | 996211 |    2,625.76 |
| High Priority   | 998365 |    1,248.00 |
| Homepage Banner | 998365 |    4,140.00 |

### Output Excel

| Invoice No       | PDF File                                        | HSN Ending 11 | HSN Not Ending 11 | Subtotal |
| ---------------- | ----------------------------------------------- | ------------: | ----------------: | -------: |
| 260513FS09005504 | Tax_Invoice_30328_13052026_260513FS09005504.pdf |      31509.19 |          21080.00 | 52589.19 |

---

## Requirements

### Python Version

Python 3.9 or higher is recommended.

Check your version:

```bash
python --version
```

---

## Dependencies

The project uses the following libraries:

### pdfplumber

Used for extracting text from PDF invoices.

### pandas

Used for creating and exporting tabular data.

### openpyxl

Used by pandas to generate Excel files.

---

## Installation

Clone the repository:

```bash
git clone https://github.com/your-username/swiggy-invoice-hsn-extractor.git
cd swiggy-invoice-hsn-extractor
```

Install dependencies:

```bash
pip install pdfplumber pandas openpyxl
```

Alternatively:

```bash
python -m pip install pdfplumber pandas openpyxl
```

---

## Project Structure

```text
swiggy-invoice-hsn-extractor/
│
├── BaseAmt.py
├── README.md
│
├── Tax_Invoice_1.pdf
├── Tax_Invoice_2.pdf
├── Tax_Invoice_3.pdf
│
└── hsn_summary.xlsx
```

Place all PDF invoices in the same folder as the script before running.

---

## Usage

Run the script:

```bash
python BaseAmt.py
```

or

```bash
py BaseAmt.py
```

---

## Output

After successful execution:

```text
hsn_summary.xlsx
```

will be created in the current directory.

The Excel file contains:

* Invoice No
* PDF File Name
* HSN Ending 11 Total
* HSN Not Ending 11 Total
* Subtotal

---

## Calculation Logic

### HSN Ending with 11

Any HSN code whose last two digits are `11` is classified into this bucket.

Example:

```text
996211
```

### HSN Not Ending with 11

All remaining HSN codes are classified into this bucket.

Example:

```text
998365
```

### Subtotal Formula

```text
Subtotal =
(HSN Ending 11 Total)
+
(HSN Not Ending 11 Total)
```

Example:

```text
31,509.19
+
21,080.00
=
52,589.19
```

---

## Error Handling

The script:

* Skips non-PDF files.
* Continues processing if a PDF fails.
* Displays the filename causing an error.
* Processes remaining invoices.

Example:

```text
Error in invoice.pdf: Unable to read PDF
```

---

## Troubleshooting

### ModuleNotFoundError: No module named 'pdfplumber'

Install dependencies:

```bash
pip install pdfplumber pandas openpyxl
```

### 'python' is not recognized

Use:

```bash
py BaseAmt.py
```

or add Python to your system PATH.

---

## Limitations

* Designed for invoices following the Swiggy invoice structure.
* Extraction accuracy depends on PDF text quality.
* Scanned-image PDFs may require OCR before processing.

---

## Future Improvements

* Direct extraction of Subtotal from invoice.
* OCR support for scanned PDFs.
* CSV export option.
* GUI interface.
* Batch processing of nested folders.
* Invoice date extraction.
* GST amount extraction.

---

## License

MIT License
