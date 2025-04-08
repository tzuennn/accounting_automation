import os
import pandas as pd
from logic.journal import filter_entries
from xlsxwriter.utility import xl_col_to_name
from utils.formatting import apply_borders

def export_report(schedule_df, journal_df, as_at, output_filename="prepayment_schedule_flexible.xlsx"):
    # Ensure output directory exists
    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)

    # Full path to file
    output_path = os.path.join(output_dir, output_filename)

    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        workbook = writer.book
        startrow = 2

        # Sheet 1: Prepayment Schedule
        schedule_df.to_excel(writer, sheet_name="Prepayment Schedule", startrow=startrow, index=False)
        ws = writer.sheets["Prepayment Schedule"]

        last_col = xl_col_to_name(len(schedule_df.columns) - 1)
        ws.merge_range(f"A1:{last_col}1", f"Prepayment schedule as at   {as_at}",
                       workbook.add_format({'bold': True, 'align': 'center'}))

        balance_idx = schedule_df.columns.get_loc("Balance")
        balance_col = xl_col_to_name(balance_idx)
        label_col = xl_col_to_name(balance_idx - 1)

        total_row = startrow + len(schedule_df) + 2
        data_start = startrow + 2

        ws.write(f"{label_col}{total_row}", "Total Balance:")
        ws.write_formula(f"{balance_col}{total_row}",
                         f'=SUM({balance_col}{data_start}:{balance_col}{total_row - 1})',
                         workbook.add_format({'bold': True, 'num_format': '#,##0.00'}))

        apply_borders(ws, schedule_df, workbook, startrow=startrow)

        # Sheet 2: Journal Entries (without "Item" column)
        journal_df_nodrop = journal_df.drop(columns="Item")
        journal_df_nodrop.to_excel(writer, sheet_name="Journal Entries", index=False)
        apply_borders(writer.sheets["Journal Entries"], journal_df_nodrop, workbook)

    print(f"Report exported to {output_path}")

def export_filtered_entries(journal_df, item_filter, month_filter):
    from logic.journal import filter_entries
    import os
    import pandas as pd
    from utils.formatting import apply_borders

    filtered = filter_entries(journal_df, item=item_filter, month_str=month_filter)
    if filtered.empty:
        print("No matching journal entries for your filter.")
        return

    # Convert item/month to filename-friendly format
    item_part = "_".join([s.replace(" ", "_") for s in item_filter]) if item_filter else "all"
    month_part = "_".join([s.replace("-", "").lower() for s in month_filter]) if month_filter else "all"
    export_name = f"entries_{item_part}_{month_part}.xlsx"
    export_path = os.path.join("output", export_name)

    os.makedirs("output", exist_ok=True)
    with pd.ExcelWriter(export_path, engine="xlsxwriter") as writer:
        filtered.to_excel(writer, sheet_name="Filtered", index=False)
        apply_borders(writer.sheets["Filtered"], filtered, writer.book)

    print(f"Exported filtered journal entries → {export_path}")
