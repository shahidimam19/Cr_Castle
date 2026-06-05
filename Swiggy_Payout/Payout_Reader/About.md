# Swiggy Payout Reader

A Python utility that consolidates multiple Swiggy Annexure Excel files into a single summary workbook.

The script:

* Reads all Swiggy invoice annexure files in a folder.
* Extracts information from the **Summary** sheet.
* Extracts and summarizes data from the **Order Level** sheet.
* Separates **Delivered** and **Cancelled** orders.
* Generates a single Excel report with one row per invoice.
* Automatically handles additional numeric columns added by Swiggy in future invoice formats.
* Applies Excel formatting for easier analysis.

---

## Features

### Summary Sheet Extraction

Extracts:

* Restaurant ID
* Payout Period
* Payout Settlement Date
* Invoice Number
* File Name

### Order Level Aggregation

For every numeric column:

* Delivered Total → `D_<Column Name>`
* Cancelled Total → `C_<Column Name>`

Example:

| D_Item Total | C_Item Total |
| ------------ | ------------ |
| 84922.00     | 426.00       |

### Excel Formatting

* Dark blue header
* White bold text
* Green Delivered columns (`D_`)
* Red Cancelled columns (`C_`)
* Auto filter
* Frozen header row
* Auto-sized columns

---

## Requirements

### Python Version

Python 3.9+

### Dependencies

Install required packages:

```bash
pip install pandas openpyxl
```

Or install from requirements.txt:

```bash
pip install -r requirements.txt
```

---

## requirements.txt

```txt
pandas
openpyxl
```

---

## Folder Structure

```text
project-folder/
│
├── Payout_Reader.py
├── invoice_Annexure_30328_15042026.xlsx
├── invoice_Annexure_30328_22042026.xlsx
├── invoice_Annexure_30328_29042026.xlsx
└── ...
```

Place the script in the same folder as the invoice files.

---

## Usage

Run:

```bash
python Payout_Reader.py
```

After completion:

```text
Consolidated_Invoice_Summary.xlsx
```

will be created in the same folder.

---

## Output Columns

### General Information

* Restaurant ID
* Payout Period
* Payout Settlement Date
* Invoice No
* File Name
* Delivered Orders
* Cancelled Orders

### Aggregated Values

For every numeric column found in the invoice:

```text
D_Item Total
C_Item Total

D_Packaging Charges
C_Packaging Charges

D_Commission
C_Commission

D_Net Payout
C_Net Payout
```

---

## Dynamic Column Handling

The script automatically detects numeric columns.

If Swiggy introduces new fields such as:

```text
Delivery Fee
Platform Fee
GST Adjustment
```

the output will automatically include:

```text
D_Delivery Fee
C_Delivery Fee

D_Platform Fee
C_Platform Fee

D_GST Adjustment
C_GST Adjustment
```

without any code changes.

---

## Ignored Files

Excel temporary files are automatically skipped:

```text
~$invoice_Annexure_xxx.xlsx
```

---

## Example Output

| Restaurant ID | Payout Period   | Invoice No | D_Item Total | C_Item Total |
| ------------- | --------------- | ---------- | ------------ | ------------ |
| 30328         | 05 Apr - 11 Apr | 30328      | 84922.00     | 426.00       |
| 30328         | 12 Apr - 18 Apr | 30328      | 92150.00     | 0.00         |

---

## License

SHAHID License

Feel free to modify and use for internal reporting and payout reconciliation.
