import os
import pandas as pd
from config import AS_AT_MONTH, START_PERIOD, END_PERIOD, ITEM_FILTER, MONTH_FILTER, INPUT_FILE
from logic.schedule import generate_schedule
from logic.journal import generate_all_entries
from logic.export import export_report, export_filtered_entries

if __name__ == "__main__":
    input_df = pd.read_csv(INPUT_FILE)

    schedules = [
        generate_schedule(
            item_name=row["Item"],
            invoice_number=int(row["Invoice Number"]),
            cost=float(row["Invoice Amount"]),
            duration_months=int(row["Duration Months"]),
            start_month_str=row["Start Month"],
            as_at_month=AS_AT_MONTH,
            start_period=START_PERIOD,
            end_period=END_PERIOD,
        )
        for _, row in input_df.iterrows()
    ]

    schedule_df = pd.concat(schedules, ignore_index=True)
    journal_df = generate_all_entries(schedule_df)

    # Export full report to output/
    export_report(schedule_df, journal_df, AS_AT_MONTH)

    # Export filtered journal to output/
    export_filtered_entries(journal_df, ITEM_FILTER, MONTH_FILTER)
