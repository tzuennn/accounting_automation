""" 
Configuration File for Prepayment Generator

Edit the fields below as needed.
You may specify multiple items or months using commas.
Leave ITEM_FILTER or MONTH_FILTER empty to include all.
"""

# === Editable Configuration ===
AS_AT_MONTH = "Oct-24"        # e.g., "Oct-24"
START_PERIOD = "Jan-24"       # e.g., "Jan-24"
END_PERIOD = "Feb-25"         # e.g., "Dec-25"

# Use comma-separated values for multiple filters
ITEM_FILTER = ["Webhosting", "Insurance"]    # e.g., ["Webhosting", "Software"]
MONTH_FILTER = ["May-24", "Jun-24"]          # e.g., ["May-24"]

INPUT_FILE = "prepaid_items.csv"  # Path to input CSV