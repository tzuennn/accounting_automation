import pandas as pd
from datetime import datetime
import calendar
from xlsxwriter.utility import xl_col_to_name

# === CONFIG ===
AS_AT_MONTH = "Oct-24"
START_PERIOD = "Jan-24"
END_PERIOD = "Dec-25"
TARGET_EXPORT_MONTH = "Jun-24"  # ← CHANGE THIS to any valid month, e.g., "Feb-25"

def generate_schedule(item_name, invoice_number, cost, duration_months, start_month_str):
    monthly_value = round(-cost / 12, 7)
    start_date = datetime.strptime("01-" + start_month_str, "%d-%b-%y")
    as_at_date = datetime.strptime("01-" + AS_AT_MONTH, "%d-%b-%y")

    schedule = {
        "Item": item_name,
        "Invoice Number": invoice_number,
        "Invoice Amount": cost,
    }

    all_months = pd.date_range(
        start=datetime.strptime("01-" + START_PERIOD, "%d-%b-%y"),
        end=datetime.strptime("01-" + END_PERIOD, "%d-%b-%y"),
        freq='MS'
    )

    filled_months = 0
    deducted_months = 0
    for m in all_months:
        col = m.strftime("%b-%y")
        if m >= start_date and filled_months < 12:
            schedule[col] = monthly_value
            filled_months += 1
            if m <= as_at_date:
                deducted_months += 1
        else:
            schedule[col] = ""

    schedule["Balance"] = round(cost - abs(monthly_value) * deducted_months, 7)
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
            if "-" in month and isinstance(row[month], (int, float)) and row[month] != 0:
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

def get_entries_for_month(journal_df, month_str):
    target_date = get_last_day(month_str)
    return journal_df[journal_df["Date"] == target_date]

# ========== MAIN ==========
if __name__ == "__main__":
    input_df = pd.read_csv("prepaid_items.csv")

    schedules = []
    for _, row in input_df.iterrows():
        schedule = generate_schedule(
            item_name=row["Item"],
            invoice_number=int(row["Invoice Number"]),
            cost=float(row["Invoice Amount"]),
            duration_months=int(row["Duration Months"]),
            start_month_str=row["Start Month"]
        )
        schedules.append(schedule)

    full_schedule = pd.concat(schedules, ignore_index=True)
    journal_df = generate_all_entries(full_schedule)

    # === Export to Excel (main schedule file) ===
    with pd.ExcelWriter("prepayment_schedule_flexible.xlsx", engine="xlsxwriter") as writer:
        workbook = writer.book
        startrow = 2

        # Sheet 1: Prepayment Schedule
        full_schedule.to_excel(writer, sheet_name="Prepayment Schedule", startrow=startrow, index=False)
        worksheet = writer.sheets["Prepayment Schedule"]

        # Merge header
        last_col_letter = xl_col_to_name(len(full_schedule.columns) - 1)
        worksheet.merge_range(f"A1:{last_col_letter}1", f"Prepayment schedule as at   {AS_AT_MONTH}",
                              workbook.add_format({'bold': True, 'align': 'center'}))

        # Total Balance
        balance_col_idx = full_schedule.columns.get_loc("Balance")
        balance_col_letter = xl_col_to_name(balance_col_idx)
        label_col_letter = xl_col_to_name(balance_col_idx - 1)

        num_data_rows = len(full_schedule)
        formula_row = startrow + num_data_rows + 2
        data_start_row = startrow + 2

        worksheet.write(f"{label_col_letter}{formula_row}", "Total Balance:")
        worksheet.write_formula(f"{balance_col_letter}{formula_row}",
                                f'=SUM({balance_col_letter}{data_start_row}:{balance_col_letter}{formula_row - 1})',
                                workbook.add_format({'bold': True, 'num_format': '#,##0.00'}))

        # Sheet 2: Journal Entries
        journal_df.to_excel(writer, sheet_name="Journal Entries", index=False)

    # === Export selected month journal entries to a dedicated Excel file ===
    sample_entries = get_entries_for_month(journal_df, TARGET_EXPORT_MONTH)

    if not sample_entries.empty:
        safe_month = TARGET_EXPORT_MONTH.replace("/", "-").replace(" ", "")
        output_excel = f"output_{safe_month}.xlsx"
        with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
            sample_entries.to_excel(writer, sheet_name="Entries", index=False)
        print(f"📁 Exported journal entries for {TARGET_EXPORT_MONTH} → {output_excel}")
    else:
        print(f"⚠️ No entries found for {TARGET_EXPORT_MONTH}.")
