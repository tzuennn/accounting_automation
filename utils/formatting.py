# utils/formatting.py
import calendar
from datetime import datetime

def get_last_day(month_str):
    dt = datetime.strptime("01-" + month_str, "%d-%b-%y")
    return dt.replace(day=calendar.monthrange(dt.year, dt.month)[1]).strftime("%d/%m/%Y")

def apply_borders(ws, df, workbook, startrow=0, startcol=0):
    border_fmt = workbook.add_format({'border': 1})

    # Track max width per column
    col_widths = [len(str(col)) for col in df.columns]

    for row in range(df.shape[0] + 1):
        for col in range(df.shape[1]):
            value = df.columns[col] if row == 0 else df.iloc[row - 1, col]
            value_str = str(value)
            ws.write(startrow + row, startcol + col, value, border_fmt)

            # Update max width
            if len(value_str) > col_widths[col]:
                col_widths[col] = len(value_str)

    # Set column widths
    for col_idx, width in enumerate(col_widths):
        ws.set_column(col_idx, col_idx, width + 2)  # +2 for padding
