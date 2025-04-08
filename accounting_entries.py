import pandas as pd
from datetime import datetime
import calendar
from xlsxwriter.utility import xl_col_to_name

# === CONFIG ===
AS_AT_MONTH = "Oct-24"
START_PERIOD = "Jan-24"
END_PERIOD = "Dec-25"
ITEM_TO_CHECK = ""         # e.g., "Webhosting", or leave "" for all
MONTH_TO_CHECK = "May-24"  # e.g., "May-24", or leave "" for all

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
    return dt.replace(day=calendar.monthrange(dt.year, dt.month)[1]).strftime("%d/%m/%Y")

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
                    "Item": item,
                    "Date": entry_date,
                    "Description": f"Prepayment amortisation for {item}",
                    "Reference": ref,
                    "Account": exp_acc,
                    "Amount": amount
                })
                entries.append({
                    "Item": item,
                    "Date": entry_date,
                    "Description": f"Prepayment amortisation for {item}",
                    "Reference": ref,
                    "Account": pre_acc,
                    "Amount": -amount
                })
    return pd.DataFrame(entries)

def filter_entries(journal_df, item="", month_str=""):
    if item:
        journal_df = journal_df[journal_df["Item"].str.lower() == item.lower()]
    if month_str:
        date_to_check = get_last_day(month_str)
        journal_df = journal_df[journal_df["Date"] == date_to_check]
    return journal_df.drop(columns=["Item"])

def apply_borders(ws, df, workbook, startrow=0, startcol=0):
    border_fmt = workbook.add_format({'border': 1})
    for row in range(df.shape[0] + 1):  # +1 for header
        for col in range(df.shape[1]):
            value = df.columns[col] if row == 0 else df.iloc[row - 1, col]
            ws.write(startrow + row, startcol + col, value, border_fmt)

# ========== MAIN ==========
if __name__ == "__main__":
    input_df = pd.read_csv("prepaid_items.csv")

    schedules = [generate_schedule(
        item_name=row["Item"],
        invoice_number=int(row["Invoice Number"]),
        cost=float(row["Invoice Amount"]),
        duration_months=int(row["Duration Months"]),
        start_month_str=row["Start Month"]
    ) for _, row in input_df.iterrows()]

    full_schedule = pd.concat(schedules, ignore_index=True)
    journal_df = generate_all_entries(full_schedule)

    # === Export full report to Excel ===
    with pd.ExcelWriter("prepayment_schedule_flexible.xlsx", engine="xlsxwriter") as writer:
        workbook = writer.book
        startrow = 2

        # Sheet 1: Prepayment Schedule
        full_schedule.to_excel(writer, sheet_name="Prepayment Schedule", startrow=startrow, index=False)
        schedule_ws = writer.sheets["Prepayment Schedule"]

        last_col_letter = xl_col_to_name(len(full_schedule.columns) - 1)
        schedule_ws.merge_range(f"A1:{last_col_letter}1", f"Prepayment schedule as at   {AS_AT_MONTH}",
                                workbook.add_format({'bold': True, 'align': 'center'}))

        balance_idx = full_schedule.columns.get_loc("Balance")
        balance_col_letter = xl_col_to_name(balance_idx)
        label_col_letter = xl_col_to_name(balance_idx - 1)

        num_rows = len(full_schedule)
        formula_row = startrow + num_rows + 2
        data_start_row = startrow + 2

        schedule_ws.write(f"{label_col_letter}{formula_row}", "Total Balance:")
        schedule_ws.write_formula(f"{balance_col_letter}{formula_row}",
                                  f'=SUM({balance_col_letter}{data_start_row}:{balance_col_letter}{formula_row - 1})',
                                  workbook.add_format({'bold': True, 'num_format': '#,##0.00'}))

        apply_borders(schedule_ws, full_schedule, workbook, startrow=startrow)

        # Sheet 2: Journal Entries
        journal_df_nodrop = journal_df.drop(columns="Item")
        journal_df_nodrop.to_excel(writer, sheet_name="Journal Entries", index=False)
        journal_ws = writer.sheets["Journal Entries"]
        apply_borders(journal_ws, journal_df_nodrop, workbook)

    # === Export filtered to separate Excel with borders ===
    filtered = filter_entries(journal_df, item=ITEM_TO_CHECK, month_str=MONTH_TO_CHECK)
    if not filtered.empty:
        item_part = ITEM_TO_CHECK.replace(" ", "_") if ITEM_TO_CHECK else "all"
        month_part = MONTH_TO_CHECK.replace("-", "").lower() if MONTH_TO_CHECK else "all"
        export_name = f"entries_{item_part}_{month_part}.xlsx"

        with pd.ExcelWriter(export_name, engine="xlsxwriter") as writer:
            filtered.to_excel(writer, sheet_name="Filtered", index=False)
            ws = writer.sheets["Filtered"]
            apply_borders(ws, filtered, writer.book)

        print(f"📁 Exported filtered journal entries → {export_name}")
    else:
        print("⚠️ No matching journal entries for your filter.")
