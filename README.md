# Prepayment Accounting Entry Generator

This Python-based solution automates the generation of prepayment schedules and their corresponding accounting entries. It helps businesses properly amortize prepaid expenses (e.g., insurance, webhosting, software) across months.

---

## Overview

When a company pays upfront for services that span multiple months, the expense must be recognized over time. This tool:

- Builds a **monthly amortization schedule** for each prepaid item.
- Generates **accounting entries** for each relevant month-end.
- Allows filtering by **item name** and/or **month**.
- Exports output in **excel** format

---

## 📁 Folder Structure

```
accounting_automation/
├── config.py               # Configuration file — easy to edit
├── main.py                 # Main entry point
├── prepaid_items.csv       # Input file with prepayment data
├── output/                 # All generated Excel reports
│
├── logic/
│   ├── schedule.py         # Generates prepayment schedules
│   ├── journal.py          # Generates and filters accounting entries
│   └── export.py           # Handles Excel report export
│
└── utils/
    └── formatting.py       # Formatting helpers for Excel (borders, etc.)
```

---

## ✅ How to Use

### 1. Prepare your data

Create a CSV file named `prepaid_items.csv` with the following structure:

```csv
Item,Invoice Number,Invoice Amount,Duration Months,Start Month
Webhosting,46248,10000,12,Jan-24
Insurance,89017,1200,12,Apr-24
Software,34567,3600,12,Mar-24
Hosting,12345,2400,12,Jun-24
```

> You may rename the file, but make sure to update `config.py`.

---

### 2. Configure the script

Open `config.py` and edit the fields:

```python
AS_AT_MONTH = "Oct-24"        # Used to calculate remaining balance
START_PERIOD = "Jan-24"       # Start of the schedule
END_PERIOD = "Dec-25"         # End of the schedule

ITEM_FILTER = "Webhosting"    # e.g., "Webhosting" or leave "" for all items
MONTH_FILTER = "May-24"       # e.g., "May-24" or leave "" for all months

INPUT_FILE = "prepaid_items.csv"
```

- Leave `ITEM_FILTER` or `MONTH_FILTER` as `""` to include all.

For example, if I want to retrieve all the accounting entries of May-24, set `ITEM_FILTER = ""` and `MONTH_FILTER = "May-24"`.
Similarly, if I want to retrieve all accounting entries of a particular item/service, eg Insurance, set `ITEM_FILTER = "Insurance"` and `MONTH_FILTER = ""`.

---

### 3. Run the script

From your terminal, run:

```bash
python main.py
```

This will create:

- `output/prepayment_schedule_flexible.xlsx`  
  (with full amortization schedule and accounting entries)

- `output/entries_<item>_<month>.xlsx`  
  (filtered based on your `config.py`)

---

## 🔍 Example Outputs

### Filter: **Item = Webhosting**, **Month = May-24**

| Date       | Description                          | Reference | Account | Amount   |
|------------|--------------------------------------|-----------|---------|----------|
| 31/05/2024 | Prepayment amortisation for Webhosting | 46248     | EXP001  | 833.33   |
| 31/05/2024 | Prepayment amortisation for Webhosting | 46248     | PRE001  | -833.33  |

---

### Filter: **Month = Jun-24** (All items)

| Date       | Description                            | Reference | Account | Amount   |
|------------|----------------------------------------|-----------|---------|----------|
| 30/06/2024 | Prepayment amortisation for Webhosting | 46248     | EXP001  | 833.33   |
| 30/06/2024 | Prepayment amortisation for Webhosting | 46248     | PRE001  | -833.33  |
| 30/06/2024 | Prepayment amortisation for Insurance  | 89017     | EXP002  | 100.00   |
| 30/06/2024 | Prepayment amortisation for Insurance  | 89017     | PRE002  | -100.00  |
| 30/06/2024 | Prepayment amortisation for Software   | 34567     | EXP003  | 300.00   |
| 30/06/2024 | Prepayment amortisation for Software   | 34567     | PRE003  | -300.00  |
| 30/06/2024 | Prepayment amortisation for Hosting    | 12345     | EXP004  | 200.00   |
| 30/06/2024 | Prepayment amortisation for Hosting    | 12345     | PRE004  | -200.00  |

---

### Filter: **Item = Insurance (All months)**

| Date       | Description                          | Reference | Account | Amount   |
|------------|--------------------------------------|-----------|---------|----------|
| 30/04/2024 | Prepayment amortisation for Insurance | 89017     | EXP002  | 100.00   |
| 30/04/2024 | Prepayment amortisation for Insurance | 89017     | PRE002  | -100.00  |
| 31/05/2024 | Prepayment amortisation for Insurance | 89017     | EXP002  | 100.00   |
| 31/05/2024 | Prepayment amortisation for Insurance | 89017     | PRE002  | -100.00  |
| 30/06/2024 | Prepayment amortisation for Insurance | 89017     | EXP002  | 100.00   |
| 30/06/2024 | Prepayment amortisation for Insurance | 89017     | PRE002  | -100.00  |
| ...        | ...                                  | ...       | ...     | ...      |
| 31/03/2025 | Prepayment amortisation for Insurance | 89017     | PRE002  | -100.00  |

---

## Notes

- Only monthly amortization is supported.
- EXP/PRE account codes are generated sequentially (e.g., EXP001, PRE001).

---

## Requirements

- Python 3.8+
- Libraries: `pandas`, `xlsxwriter`

Install with:

```bash
pip install -r requirements.txt
```

---
