# config.py
"""
Configuration File for Prepayment Generator

Edit the fields below as needed.
Leave ITEM_FILTER or MONTH_FILTER blank ("") to include all.
"""
# === Editable Configuration ===
AS_AT_MONTH = "Oct-24"        # e.g., "Oct-24"
START_PERIOD = "Jan-24"       # e.g., "Jan-24"
END_PERIOD = "Dec-25"         # e.g., "Dec-25"
ITEM_FILTER = ""              # e.g., "Webhosting" or leave "" for all items
MONTH_FILTER = "Jun-24"       # e.g., "May-24" or leave "" for all months
INPUT_FILE = "prepaid_items.csv"  # Path to input CSV
