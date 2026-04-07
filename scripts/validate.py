"""
Validate uploaded Excel files against expected schemas.
"""

import os
import pandas as pd
from scripts.config import SCHEMAS, get_required_columns, get_column_types, DEALERS_CSV


def normalize_columns(df):
    """Normalize DataFrame column names: lowercase, strip, replace spaces with underscores."""
    df.columns = [c.lower().strip().replace(" ", "_") for c in df.columns]
    return df


def validate_upload(file_path, dataset_type):
    """
    Validate an uploaded Excel file against the expected schema.

    Args:
        file_path: Path to the Excel file.
        dataset_type: One of 'dealers', 'install_events', 'usage_metrics'.

    Returns:
        (is_valid, messages): Tuple of bool and list of error/warning strings.
    """
    messages = []

    if dataset_type not in SCHEMAS:
        return False, [f"Unknown dataset type: '{dataset_type}'. Must be one of: {', '.join(SCHEMAS.keys())}"]

    if not os.path.exists(file_path):
        return False, [f"File not found: {file_path}"]

    # Read the Excel file
    try:
        df = pd.read_excel(file_path, engine="openpyxl")
    except Exception as e:
        return False, [f"Could not read Excel file: {e}"]

    if df.empty:
        return False, ["The uploaded file is empty (no data rows)."]

    # Normalize column names
    df = normalize_columns(df)

    schema = SCHEMAS[dataset_type]
    col_types = get_column_types(dataset_type)
    required = get_required_columns(dataset_type)

    # Check required columns
    missing = [c for c in required if c not in df.columns]
    if missing:
        messages.append(f"ERROR: Missing required columns: {', '.join(missing)}")
        return False, messages

    # Warn about unexpected columns
    known_cols = {col[0] for col in schema["columns"]}
    extra = [c for c in df.columns if c not in known_cols]
    if extra:
        messages.append(f"WARNING: Unexpected columns (will be ignored): {', '.join(extra)}")

    # Check for blank required fields
    for col in required:
        null_count = df[col].isna().sum()
        if null_count > 0:
            messages.append(f"ERROR: Column '{col}' has {null_count} blank value(s). This field is required.")

    # Validate data types
    for col_name, col_type in col_types.items():
        if col_name not in df.columns:
            continue
        if col_type == "date":
            try:
                pd.to_datetime(df[col_name].dropna(), format="mixed")
            except Exception:
                messages.append(f"ERROR: Column '{col_name}' contains invalid date values.")
        elif col_type == "integer":
            non_null = df[col_name].dropna()
            if len(non_null) > 0:
                try:
                    vals = pd.to_numeric(non_null, errors="coerce")
                    bad_count = vals.isna().sum()
                    if bad_count > 0:
                        messages.append(f"ERROR: Column '{col_name}' has {bad_count} non-numeric value(s).")
                except Exception:
                    messages.append(f"ERROR: Column '{col_name}' contains invalid integer values.")
        elif col_type == "float":
            non_null = df[col_name].dropna()
            if len(non_null) > 0:
                try:
                    vals = pd.to_numeric(non_null, errors="coerce")
                    bad_count = vals.isna().sum()
                    if bad_count > 0:
                        messages.append(f"ERROR: Column '{col_name}' has {bad_count} non-numeric value(s).")
                except Exception:
                    messages.append(f"ERROR: Column '{col_name}' contains invalid float values.")

    # Validate event_type values for install_events
    if dataset_type == "install_events" and "event_type" in df.columns:
        valid_types = schema.get("valid_event_types", [])
        if valid_types:
            invalid = df[~df["event_type"].str.lower().isin(valid_types)]["event_type"].unique()
            if len(invalid) > 0:
                messages.append(
                    f"ERROR: Column 'event_type' has invalid values: {', '.join(str(v) for v in invalid)}. "
                    f"Must be one of: {', '.join(valid_types)}"
                )

    # Check for duplicate primary keys
    pk = schema["primary_key"]
    pk_present = [c for c in pk if c in df.columns]
    if len(pk_present) == len(pk):
        dupes = df.duplicated(subset=pk, keep=False)
        dupe_count = dupes.sum()
        if dupe_count > 0:
            messages.append(f"WARNING: {dupe_count} duplicate rows found based on key columns: {', '.join(pk)}")

    # Referential integrity: check dealer_id exists in master dealers list
    if dataset_type != "dealers" and "dealer_id" in df.columns and os.path.exists(DEALERS_CSV):
        try:
            dealers_df = pd.read_csv(DEALERS_CSV, dtype=str)
            known_ids = set(dealers_df["dealer_id"].values)
            uploaded_ids = set(df["dealer_id"].dropna().astype(str).values)
            unknown = uploaded_ids - known_ids
            if unknown:
                messages.append(
                    f"WARNING: {len(unknown)} dealer ID(s) not found in master dealer list: "
                    f"{', '.join(sorted(unknown)[:10])}"
                    + (" ..." if len(unknown) > 10 else "")
                )
        except Exception:
            pass  # Don't fail validation if we can't read dealers file

    # Determine overall validity
    has_errors = any(msg.startswith("ERROR:") for msg in messages)
    if not has_errors and not messages:
        messages.append("Validation passed. No issues found.")

    return not has_errors, messages
