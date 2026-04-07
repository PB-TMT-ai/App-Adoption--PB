"""
Central configuration for the App Adoption Tracking System.
All paths, column schemas, and constants are defined here.
"""

import os

# --- Paths ---
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_DIR = os.path.join(BASE_DIR, "templates")
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
PROCESSED_DIR = os.path.join(UPLOAD_DIR, "processed")
DATA_DIR = os.path.join(BASE_DIR, "data")
REPORT_DIR = os.path.join(BASE_DIR, "reports")

# Master CSV file path
APP_ADOPTION_CSV = os.path.join(DATA_DIR, "app_adoption.csv")

# --- Valid Status Values ---
VALID_STATUSES = ["Installed", "Not Installed", "Uninstalled", "Tech Issue", "Not Working"]

# --- Column Schema ---
# Maps original Excel column names to (internal_name, type, required)
COLUMNS = [
    ("Account Sf ID",              "account_sf_id",          "string",  True),
    ("Name of the Dealer",         "dealer_name",            "string",  True),
    ("State",                      "state",                  "string",  False),
    ("Zone",                       "zone",                   "string",  False),
    ("Distributor Name",           "distributor_name",       "string",  False),
    ("Account Owner As per SF",    "account_owner",          "string",  False),
    ("GST No.",                    "gst_no",                 "string",  False),
    ("Mobile No.",                 "mobile_no",              "string",  False),
    ("Dealer classification",      "dealer_classification",  "string",  False),
    ("# orders (total)",           "orders_total",           "integer", False),
    ("# orders (via. app.)",       "orders_via_app",         "integer", False),
    ("% orders via. app.",         "pct_orders_via_app",     "float",   False),
    ("Order quantity (total)",     "order_qty_total",        "numeric", False),
    ("Quantity (via. app.)",       "qty_via_app",            "numeric", False),
    ("% order quantity (via. app.)", "pct_qty_via_app",      "float",   False),
    ("Status",                     "status",                 "string",  True),
]

PRIMARY_KEY = "Account Sf ID"

# --- Helper Functions ---

def get_excel_columns():
    """Return list of original Excel column names."""
    return [col[0] for col in COLUMNS]


def get_required_columns():
    """Return list of required original Excel column names."""
    return [col[0] for col in COLUMNS if col[3]]


def get_column_type_map():
    """Return dict of original_column_name -> type_string."""
    return {col[0]: col[2] for col in COLUMNS}


def get_column_rename_map():
    """Return dict of original_column_name -> internal_name."""
    return {col[0]: col[1] for col in COLUMNS}


# --- Sample Data (for template) ---
SAMPLE_DATA = [
    {
        "Account Sf ID": "0015g00001EXAMPLE",
        "Name of the Dealer": "Example Dealer Name",
        "State": "UTTAR PRADESH",
        "Zone": "North",
        "Distributor Name": "Example Distributor Pvt. Ltd.",
        "Account Owner As per SF": "John Doe",
        "GST No.": "09AAAAA0000A1Z5",
        "Mobile No.": "9876543210",
        "Dealer classification": "Active",
        "# orders (total)": 10,
        "# orders (via. app.)": 5,
        "% orders via. app.": 50.0,
        "Order quantity (total)": 100,
        "Quantity (via. app.)": 50,
        "% order quantity (via. app.)": 50.0,
        "Status": "Installed",
    },
    {
        "Account Sf ID": "0015g00002EXAMPLE",
        "Name of the Dealer": "Another Dealer Name",
        "State": "HARYANA",
        "Zone": "North",
        "Distributor Name": "Another Distributor",
        "Account Owner As per SF": "Jane Smith",
        "GST No.": "06BBBBB0000B1Z5",
        "Mobile No.": "9123456789",
        "Dealer classification": "Active",
        "# orders (total)": 0,
        "# orders (via. app.)": 0,
        "% orders via. app.": "",
        "Order quantity (total)": "-",
        "Quantity (via. app.)": "-",
        "% order quantity (via. app.)": "",
        "Status": "Not Installed",
    },
]
