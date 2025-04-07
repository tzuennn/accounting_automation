import pandas as pd
from datetime import datetime
import calendar
import xlsxwriter

AS_AT_MONTH = "Oct-24"
YEAR = "2024"

def generate_schedule(item_name, invoice_number, cost, duration_months, start_month_str):
    monthly_value = round(-cost / 12, 7)  
    start_date = datetime.strptime("01-" + start_month_str, "%d-%b-%y")
    as_at_date = datetime.strptime("01-" + AS_AT_MONTH, "%d-%b-%y")

    schedule = {
        "Item": item_name,
        "Invoice Number": invoice_number,
        "Invoice Amount": cost,
    }

    all_months = pd.date_range(f"{YEAR}-01-01", f"{YEAR}-12-01", freq='MS')
    num_deducted_months = 0

    for m in all_months:
        col = m.strftime("%b-%y")
        if m >= start_date:
            schedule[col] = monthly_value
            if m <= as_at_date:
                num_deducted_months += 1
        else:
            schedule[col] = ""

    # Corrected balance calculation: amount - (num deducted × abs(monthly))
    schedule["Balance"] = round(cost - abs(monthly_value) * num_deducted_months, 7)
    return pd.DataFrame([schedule])

def get_last_day(month_str):
    dt = datetime.strptime("01-" + month_str, "%d-%b-%y")
    last_day = calendar.monthrange(dt.year, dt.month)[1]
    return dt.replace(day=last_day).strftime("%d/%m/%Y")

def generate_all_entries(schedule_df):
    entries = []
    for idx, row in schedule_df.iterrows():
        item = row["Item"]
        ref = row["Invoice Number"]
        exp_acc = f"EXP{str(idx+1).zfill(3)}"
        pre_acc = f"PRE{str(idx+1).zfill(3)}"

        for month in row.index:
            if month.endswith("-24") and isinstance(row[month], (int, float)) and row[month] != 0:
                amount = round(abs(row[month]), 2)
                entry_date = get_last_day(month)

                entries.append({
                    "Date": entry_date,
                    "Description": f"Prepayment amortisation for {item}",
                    "Reference": ref,
                    "Account": exp_acc,
                    "Amount": amount
                })
                entries.append({
                    "Date": entry_date,
                    "Description": f"Prepayment amortisation for {item}",
                    "Reference": ref,
                    "Account": pre_acc,
                    "Amount": -amount
                })

    return pd.DataFrame(entries)

if __name__ == "__main__":
    webhosting = generate_schedule("Webhosting", "46248", 10000, 12, "Jan-24")
    insurance = generate_schedule("Insurance", "89017", 1200, 12, "Apr-24")

    full_schedule = pd.concat([webhosting, insurance], ignore_index=True)
    journal_df = generate_all_entries(full_schedule)

    # === Export to Excel ===
    with pd.ExcelWriter("prepayment_schedule_oct24.xlsx", engine="xlsxwriter") as writer:
        workbook = writer.book
        startrow = 2

        # --- Sheet 1: Prepayment Schedule ---
        full_schedule.to_excel(writer, sheet_name="Prepayment Schedule", startrow=startrow, index=False)
        worksheet = writer.sheets["Prepayment Schedule"]

        # Merge header title
        last_col_letter = chr(65 + len(full_schedule.columns) - 1)
        worksheet.merge_range(f'A1:{last_col_letter}1', f"Prepayment schedule as at   {AS_AT_MONTH}",
                              workbook.add_format({'bold': True, 'align': 'center'}))

        # Total Balance
        balance_col_idx = full_schedule.columns.get_loc("Balance")
        balance_col_letter = chr(65 + balance_col_idx)
        total_row_excel = startrow + 1 + len(full_schedule)  # data starts at row 3

        worksheet.write(f"{balance_col_letter[:-1]}{total_row_excel}", "Total Balance:")
        worksheet.write(f"{balance_col_letter}{total_row_excel}",
                        f'=SUM({balance_col_letter}{startrow + 3}:{balance_col_letter}{total_row_excel - 1})',
                        workbook.add_format({'bold': True, 'num_format': '#,##0.00'}))

        # --- Sheet 2: Journal Entries ---
        journal_df.to_excel(writer, sheet_name="Journal Entries", index=False)

    print("✅ Exported to prepayment_schedule_oct24.xlsx with correct spreads and balance total.")
