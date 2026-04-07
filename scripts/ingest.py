"""
Ingest uploaded Excel files into the master CSV dataset.
"""

import os
import shutil
from datetime import datetime

import pandas as pd

from scripts.config import (
    get_excel_columns, APP_ADOPTION_CSV, DATA_DIR, PROCESSED_DIR, PRIMARY_KEY,
)
from scripts.validate import validate_upload


def _clean_numeric(series):
    """Convert a series with '-' placeholders and mixed types to numeric."""
    return pd.to_numeric(
        series.astype(str).str.strip().replace("-", "0").replace("", "0"),
        errors="coerce",
    ).fillna(0)


def ingest_file(file_path):
    """
    Validate and merge an uploaded Excel file into the master CSV.

    Args:
        file_path: Path to the uploaded Excel file.

    Returns:
        (rows_affected, messages): Tuple of int and list of status strings.

    Raises:
        ValueError: If validation fails.
    """
    is_valid, val_messages = validate_upload(file_path)
    if not is_valid:
        raise ValueError("Validation failed:\n" + "\n".join(val_messages))

    # Read upload
    upload_df = pd.read_excel(file_path, engine="openpyxl", dtype=str)
    upload_df.columns = [c.strip() for c in upload_df.columns]

    # Keep only known columns
    known_cols = get_excel_columns()
    cols_to_keep = [c for c in known_cols if c in upload_df.columns]
    upload_df = upload_df[cols_to_keep]

    # Clean numeric columns
    numeric_cols = ["# orders (total)", "# orders (via. app.)",
                    "Order quantity (total)", "Quantity (via. app.)"]
    for col in numeric_cols:
        if col in upload_df.columns:
            upload_df[col] = _clean_numeric(upload_df[col])

    float_cols = ["% orders via. app.", "% order quantity (via. app.)"]
    for col in float_cols:
        if col in upload_df.columns:
            upload_df[col] = pd.to_numeric(
                upload_df[col].astype(str).str.strip().replace("", None),
                errors="coerce",
            )

    # Read existing master CSV or start empty
    os.makedirs(DATA_DIR, exist_ok=True)
    if os.path.exists(APP_ADOPTION_CSV) and os.path.getsize(APP_ADOPTION_CSV) > 0:
        master_df = pd.read_csv(APP_ADOPTION_CSV, dtype=str)
        # Re-parse numeric columns
        for col in numeric_cols:
            if col in master_df.columns:
                master_df[col] = _clean_numeric(master_df[col])
        for col in float_cols:
            if col in master_df.columns:
                master_df[col] = pd.to_numeric(master_df[col], errors="coerce")
    else:
        master_df = pd.DataFrame(columns=cols_to_keep)

    messages = list(val_messages)

    # Upsert on primary key
    pk = PRIMARY_KEY
    if master_df.empty:
        merged = upload_df.copy()
        new_count = len(upload_df)
        updated_count = 0
    else:
        existing_keys = set(master_df[pk].astype(str))
        upload_keys = upload_df[pk].astype(str)

        new_mask = ~upload_keys.isin(existing_keys)
        update_mask = upload_keys.isin(existing_keys)

        new_count = int(new_mask.sum())
        updated_count = int(update_mask.sum())

        # Remove old rows that will be updated
        if updated_count > 0:
            update_key_set = set(upload_keys[update_mask])
            keep_mask = ~master_df[pk].astype(str).isin(update_key_set)
            master_df = master_df[keep_mask]

        merged = pd.concat([master_df, upload_df], ignore_index=True)

    # Write back to CSV
    merged.to_csv(APP_ADOPTION_CSV, index=False)

    messages.append(f"Dealers: {new_count} new, {updated_count} updated. Total: {len(merged)} rows.")

    # Move processed file
    os.makedirs(PROCESSED_DIR, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_name = os.path.basename(file_path)
    dest = os.path.join(PROCESSED_DIR, f"{timestamp}_{base_name}")
    shutil.copy2(file_path, dest)
    messages.append(f"Copy saved to: {dest}")

    return new_count + updated_count, messages
