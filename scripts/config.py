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

# Master CSV file paths
DEALERS_CSV = os.path.join(DATA_DIR, "dealers.csv")
INSTALL_EVENTS_CSV = os.path.join(DATA_DIR, "install_events.csv")
USAGE_METRICS_CSV = os.path.join(DATA_DIR, "usage_metrics.csv")

# --- Constants ---
INACTIVE_DAYS = 30          # Days without activity to be considered inactive
AT_RISK_DAYS = 14           # Days without activity to appear on "At Risk" list
ENGAGEMENT_HIGH = 70        # Engagement score >= 70 is "High"
ENGAGEMENT_MEDIUM = 40      # Engagement score >= 40 is "Medium", below is "Low"

# --- Column Schemas ---
# Each schema defines: column name, python type string, and whether it's required.

SCHEMAS = {
    "dealers": {
        "columns": [
            ("dealer_id",      "string",  True),
            ("dealer_name",    "string",  True),
            ("dealer_group",   "string",  False),
            ("city",           "string",  False),
            ("state",          "string",  False),
            ("zip_code",       "string",  False),
            ("region",         "string",  False),
            ("contact_name",   "string",  False),
            ("contact_email",  "string",  False),
            ("contact_phone",  "string",  False),
            ("onboarded_date", "date",    False),
        ],
        "primary_key": ["dealer_id"],
        "csv_path": DEALERS_CSV,
        "template_name": "dealer_info_template.xlsx",
    },
    "install_events": {
        "columns": [
            ("dealer_id",    "string",  True),
            ("event_type",   "string",  True),
            ("event_date",   "date",    True),
            ("app_version",  "string",  False),
            ("device_os",    "string",  False),
            ("device_model", "string",  False),
            ("notes",        "string",  False),
        ],
        "primary_key": ["dealer_id", "event_type", "event_date"],
        "csv_path": INSTALL_EVENTS_CSV,
        "template_name": "install_events_template.xlsx",
        "valid_event_types": ["install", "uninstall"],
    },
    "usage_metrics": {
        "columns": [
            ("dealer_id",          "string",  True),
            ("period_start",       "date",    True),
            ("period_end",         "date",    True),
            ("login_count",        "integer", False),
            ("sessions_count",     "integer", False),
            ("actions_performed",  "integer", False),
            ("last_active_date",   "date",    False),
            ("avg_session_minutes","float",   False),
            ("features_used",      "string",  False),
            ("engagement_score",   "float",   False),
        ],
        "primary_key": ["dealer_id", "period_start", "period_end"],
        "csv_path": USAGE_METRICS_CSV,
        "template_name": "usage_metrics_template.xlsx",
    },
}

# --- Sample Data (for templates) ---
SAMPLE_DATA = {
    "dealers": [
        {
            "dealer_id": "DLR001",
            "dealer_name": "Smith Auto Group",
            "dealer_group": "Smith Holdings",
            "city": "Dallas",
            "state": "TX",
            "zip_code": "75201",
            "region": "South",
            "contact_name": "John Smith",
            "contact_email": "john@smithauto.com",
            "contact_phone": "214-555-0100",
            "onboarded_date": "2025-01-15",
        },
        {
            "dealer_id": "DLR002",
            "dealer_name": "Johnson Motors",
            "dealer_group": "",
            "city": "Chicago",
            "state": "IL",
            "zip_code": "60601",
            "region": "Midwest",
            "contact_name": "Jane Johnson",
            "contact_email": "jane@johnsonmotors.com",
            "contact_phone": "312-555-0200",
            "onboarded_date": "2025-02-01",
        },
    ],
    "install_events": [
        {
            "dealer_id": "DLR001",
            "event_type": "install",
            "event_date": "2025-01-20",
            "app_version": "2.4.1",
            "device_os": "iOS",
            "device_model": "iPhone 15",
            "notes": "",
        },
        {
            "dealer_id": "DLR002",
            "event_type": "install",
            "event_date": "2025-02-05",
            "app_version": "2.4.1",
            "device_os": "Android",
            "device_model": "Samsung Galaxy S24",
            "notes": "First install",
        },
    ],
    "usage_metrics": [
        {
            "dealer_id": "DLR001",
            "period_start": "2025-03-01",
            "period_end": "2025-03-31",
            "login_count": 22,
            "sessions_count": 30,
            "actions_performed": 145,
            "last_active_date": "2025-03-28",
            "avg_session_minutes": 8.5,
            "features_used": "inventory,messaging,analytics",
            "engagement_score": 78.0,
        },
        {
            "dealer_id": "DLR002",
            "period_start": "2025-03-01",
            "period_end": "2025-03-31",
            "login_count": 5,
            "sessions_count": 6,
            "actions_performed": 20,
            "last_active_date": "2025-03-10",
            "avg_session_minutes": 3.2,
            "features_used": "inventory",
            "engagement_score": 25.0,
        },
    ],
}


def get_column_names(dataset_type):
    """Return list of column names for a dataset type."""
    return [col[0] for col in SCHEMAS[dataset_type]["columns"]]


def get_required_columns(dataset_type):
    """Return list of required column names for a dataset type."""
    return [col[0] for col in SCHEMAS[dataset_type]["columns"] if col[2]]


def get_column_types(dataset_type):
    """Return dict of column_name -> type_string for a dataset type."""
    return {col[0]: col[1] for col in SCHEMAS[dataset_type]["columns"]}


def detect_dataset_type(columns):
    """Auto-detect dataset type by matching column headers against known schemas."""
    normalized = {c.lower().strip().replace(" ", "_") for c in columns}
    best_match = None
    best_score = 0
    for dtype, schema in SCHEMAS.items():
        schema_cols = {col[0] for col in schema["columns"]}
        required_cols = {col[0] for col in schema["columns"] if col[2]}
        # All required columns must be present
        if not required_cols.issubset(normalized):
            continue
        score = len(schema_cols & normalized) / len(schema_cols)
        if score > best_score:
            best_score = score
            best_match = dtype
    return best_match
