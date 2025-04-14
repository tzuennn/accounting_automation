# logic/schedule.py
import pandas as pd
from datetime import datetime

def generate_schedule(item_name, invoice_number, cost, duration_months, start_month_str, as_at_month, start_period, end_period):
    monthly_value = round(-cost / duration_months, 7)
    start_date = datetime.strptime("01-" + start_month_str, "%d-%b-%y")
    as_at_date = datetime.strptime("01-" + as_at_month, "%d-%b-%y")
    schedule = {
        "Item": item_name,
        "Invoice Number": invoice_number,
        "Invoice Amount": cost,
    }

    all_months = pd.date_range(
        start=datetime.strptime("01-" + start_period, "%d-%b-%y"),
        end=datetime.strptime("01-" + end_period, "%d-%b-%y"),
        freq='MS'
    )

    filled_months = 0
    deducted_months = 0
    for m in all_months:
        col = m.strftime("%b-%y")
        if m >= start_date and filled_months < duration_months:
            schedule[col] = monthly_value
            filled_months += 1
            if m <= as_at_date:
                deducted_months += 1
        else:
            schedule[col] = ""
    schedule["Balance"] = round(cost - abs(monthly_value) * deducted_months, 7)
    return pd.DataFrame([schedule])
