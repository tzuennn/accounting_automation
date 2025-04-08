from utils.formatting import get_last_day
import pandas as pd

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

def filter_entries(journal_df, item=None, month_str=None):
    if item:
        item_set = set(i.strip().lower() for i in item) if isinstance(item, list) else {item.strip().lower()}
        journal_df = journal_df[journal_df["Item"].str.lower().isin(item_set)]
    if month_str:
        dates_to_match = {get_last_day(m.strip()) for m in month_str} if isinstance(month_str, list) else {get_last_day(month_str)}
        journal_df = journal_df[journal_df["Date"].isin(dates_to_match)]
    return journal_df.drop(columns=["Item"])