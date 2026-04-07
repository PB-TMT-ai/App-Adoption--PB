"""
Validate uploaded Excel files against expected schema.
"""

import os
import pandas as pd
from scripts.config import (
    get_excel_columns, get_required_columns, get_column_type_map,
    PRIMARY_KEY, VALID_STATUSES,
)


def validate_upload(file_path):
    """
    Validate an uploaded Excel file against the expected app adoption schema.

    Args:
        file_path: Path to the Excel file.

    Returns:
        (is_valid, messages): Tuple of bool and list of error/warning strings.
    """
    messages = []

    if not os.path.exists(file_path):
        return False, [f"File not found: {file_path}"]

    try:
        df = pd.read_excel(file_path, engine="openpyxl")
    except Exception as e:
        return False, [f"Could not read Excel file: {e}"]

    if df.empty:
        return False, ["The uploaded file is empty (no data rows)."]

    # Strip whitespace from column names
    df.columns = [c.strip() for c in df.columns]

    required = get_required_columns()
    expected = get_excel_columns()
    col_types = get_column_type_map()

    # Check required columns
    missing = [c for c in required if c not in df.columns]
    if missing:
        messages.append(f"ERROR: Missing required columns: {', '.join(missing)}")
        return False, messages

    # Warn about unexpected columns
    extra = [c for c in df.columns if c not in expected]
    if extra:
        messages.append(f"WARNING: Unexpected columns (will be ignored): {', '.join(extra)}")

    # Check for blank required fields (warning for Name, error for ID/Status)
    for col in required:
        if col in df.columns:
            null_count = df[col].isna().sum()
            if null_count > 0:
                if col == "Name of the Dealer":
                    messages.append(f"WARNING: Column '{col}' has {null_count} blank value(s).")
                else:
                    messages.append(f"ERROR: Column '{col}' has {null_count} blank value(s). This field is required.")

    # Validate Status values
    if "Status" in df.columns:
        invalid_statuses = df[~df["Status"].isin(VALID_STATUSES)]["Status"].dropna().unique()
        if len(invalid_statuses) > 0:
            messages.append(
                f"WARNING: Column 'Status' has unexpected values: {', '.join(str(v) for v in invalid_statuses)}. "
                f"Expected: {', '.join(VALID_STATUSES)}"
            )

    # Validate numeric columns
    for col_name, col_type in col_types.items():
        if col_name not in df.columns:
            continue
        if col_type in ("integer", "float", "numeric"):
            non_null = df[col_name].dropna()
            # Filter out "-" placeholder values
            non_null = non_null[non_null.astype(str).str.strip() != "-"]
            if len(non_null) > 0:
                vals = pd.to_numeric(non_null, errors="coerce")
                bad_count = vals.isna().sum()
                if bad_count > 0:
                    messages.append(f"WARNING: Column '{col_name}' has {bad_count} non-numeric value(s) (excluding '-').")

    # Check for duplicate primary keys
    if PRIMARY_KEY in df.columns:
        dupes = df.duplicated(subset=[PRIMARY_KEY], keep=False)
        dupe_count = dupes.sum()
        if dupe_count > 0:
            messages.append(f"WARNING: {dupe_count} duplicate rows found based on '{PRIMARY_KEY}'")

    has_errors = any(msg.startswith("ERROR:") for msg in messages)
    if not has_errors and not messages:
        messages.append("Validation passed. No issues found.")

    return not has_errors, messages
