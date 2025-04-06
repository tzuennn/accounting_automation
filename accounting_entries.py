import pandas as pd
from datetime import datetime
import calendar

# ========== Part 1: Generate amortization schedule ==========
def generate_schedule(item_name, invoice_number, cost, duration_months, start_month_str):
    start_date = datetime.strptime("01-" + start_month_str, "%d-%b-%y")
    months = pd.date_range(start=start_date, periods=duration_months, freq='MS')
    monthly_value = round(-cost / duration_months, 4)

    schedule = {
        "Item": item_name,
        "Invoice Number": invoice_number,
        "Invoice Amount": cost,
    }

    # Fill monthly amortization values
    for m in months:
        col_name = m.strftime("%b-%y")
        schedule[col_name] = monthly_value

    # Ensure all months of 2024 are present (optional, for formatting consistency)
    for m in pd.date_range("2024-01-01", "2024-12-01", freq='MS'):
        col = m.strftime("%b-%y")
        schedule.setdefault(col, None)

    # Compute balance left over (if any rounding issue)
    schedule["Balance"] = round(cost + (monthly_value * duration_months), 4)
    return pd.DataFrame([schedule])

# ========== Part 2: Generate journal entries ==========
def get_last_day(month_str):
    dt = datetime.strptime("01-" + month_str, "%d-%b-%y")
    last_day = calendar.monthrange(dt.year, dt.month)[1]
    return dt.replace(day=last_day).strftime("%d/%m/%Y")

def generate_entries(schedule_df, month_str):
    entries = []
    entry_date = get_last_day(month_str)

    for idx, row in schedule_df.iterrows():
        amount = row.get(month_str)
        if amount and amount != 0:
            item = row["Item"]
            ref = row["Invoice Number"]
            exp_acc = f"EXP{str(idx+1).zfill(3)}"
            pre_acc = f"PRE{str(idx+1).zfill(3)}"

            entries.append({
                "Date": entry_date,
                "Description": f"Prepayment amortisation for {item}",
                "Reference": ref,
                "Account": exp_acc,
                "Amount": round(abs(amount), 2)
            })
            entries.append({
                "Date": entry_date,
                "Description": f"Prepayment amortisation for {item}",
                "Reference": ref,
                "Account": pre_acc,
                "Amount": round(-abs(amount), 2)
            })
    return pd.DataFrame(entries)

# ========== Part 3: Sample usage ==========
if __name__ == "__main__":
    # Example 1: Webhosting - 12 months from Jan-24
    webhosting = generate_schedule(
        item_name="Webhosting",
        invoice_number="46248",
        cost=10000,
        duration_months=12,
        start_month_str="Jan-24"
    )

    # Example 2: Insurance - 9 months from Apr-24
    insurance = generate_schedule(
        item_name="Insurance",
        invoice_number="89017",
        cost=1200,
        duration_months=9,
        start_month_str="Apr-24"
    )

    # Combine the schedules into one
    combined_schedule = pd.concat([webhosting, insurance], ignore_index=True)

    # === Print amortization schedule ===
    print("========== AMORTIZATION SCHEDULE ==========")
    print(combined_schedule.to_string(index=False))

    # === Generate journal entries for a specific month ===
    month_input = "May-24"
    print(f"\n========== JOURNAL ENTRIES FOR {month_input.upper()} ==========")
    journal_df = generate_entries(combined_schedule, month_input)

    if journal_df.empty:
        print(f"No amortizations found for {month_input}.")
    else:
        print(journal_df.to_string(index=False))
