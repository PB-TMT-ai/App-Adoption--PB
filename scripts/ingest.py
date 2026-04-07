"""
Ingest uploaded Excel files into master CSV datasets.
"""

import os
import shutil
from datetime import datetime

import pandas as pd

from scripts.config import SCHEMAS, get_column_names, DATA_DIR, PROCESSED_DIR
from scripts.validate import validate_upload, normalize_columns


def ingest_file(file_path, dataset_type):
    """
    Validate and merge an uploaded Excel file into the master CSV.

    Args:
        file_path: Path to the uploaded Excel file.
        dataset_type: One of 'dealers', 'install_events', 'usage_metrics'.

    Returns:
        (rows_affected, messages): Tuple of int and list of status strings.

    Raises:
        ValueError: If validation fails.
    """
    # Validate first
    is_valid, val_messages = validate_upload(file_path, dataset_type)
    if not is_valid:
        raise ValueError("Validation failed:\n" + "\n".join(val_messages))

    schema = SCHEMAS[dataset_type]
    csv_path = schema["csv_path"]
    pk = schema["primary_key"]
    known_cols = get_column_names(dataset_type)

    # Read upload
    upload_df = pd.read_excel(file_path, engine="openpyxl")
    upload_df = normalize_columns(upload_df)

    # Keep only known columns
    cols_to_keep = [c for c in known_cols if c in upload_df.columns]
    upload_df = upload_df[cols_to_keep]

    # Convert date columns
    col_types = {col[0]: col[1] for col in schema["columns"]}
    for col_name in cols_to_keep:
        if col_types.get(col_name) == "date":
            upload_df[col_name] = pd.to_datetime(upload_df[col_name], format="mixed", errors="coerce")

    # Normalize event_type to lowercase for install_events
    if dataset_type == "install_events" and "event_type" in upload_df.columns:
        upload_df["event_type"] = upload_df["event_type"].str.lower().str.strip()

    # Read existing master CSV or start empty
    os.makedirs(DATA_DIR, exist_ok=True)
    if os.path.exists(csv_path) and os.path.getsize(csv_path) > 0:
        master_df = pd.read_csv(csv_path, dtype=str)
        # Parse date columns back
        for col_name in master_df.columns:
            if col_types.get(col_name) == "date":
                master_df[col_name] = pd.to_datetime(master_df[col_name], errors="coerce")
    else:
        master_df = pd.DataFrame(columns=known_cols)

    messages = []
    original_count = len(master_df)

    if dataset_type == "dealers":
        # Upsert: update existing dealer_id rows, append new ones
        rows_affected = _upsert(master_df, upload_df, pk)
        master_df = rows_affected[0]
        new_count = rows_affected[1]
        updated_count = rows_affected[2]
        messages.append(f"Dealers: {new_count} new, {updated_count} updated.")
        rows_affected = new_count + updated_count

    elif dataset_type == "install_events":
        # Append with deduplication on primary key
        before = len(master_df)
        combined = pd.concat([master_df, upload_df], ignore_index=True)
        # Convert PK columns to string for consistent dedup
        for col in pk:
            if col != "event_date":
                combined[col] = combined[col].astype(str)
        combined = combined.drop_duplicates(subset=pk, keep="last")
        master_df = combined
        rows_affected = len(master_df) - before
        messages.append(f"Install events: {rows_affected} new event(s) added.")

    elif dataset_type == "usage_metrics":
        # Upsert on (dealer_id, period_start, period_end)
        result = _upsert(master_df, upload_df, pk)
        master_df = result[0]
        new_count = result[1]
        updated_count = result[2]
        messages.append(f"Usage metrics: {new_count} new, {updated_count} updated.")
        rows_affected = new_count + updated_count

    # Write back to CSV
    master_df.to_csv(csv_path, index=False)

    # Move processed file
    os.makedirs(PROCESSED_DIR, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_name = os.path.basename(file_path)
    dest = os.path.join(PROCESSED_DIR, f"{timestamp}_{base_name}")
    shutil.move(file_path, dest)
    messages.append(f"Processed file moved to: {dest}")

    return rows_affected, messages


def _upsert(master_df, upload_df, pk):
    """
    Upsert upload_df into master_df based on primary key columns.

    Returns:
        (merged_df, new_count, updated_count)
    """
    if master_df.empty:
        return upload_df.copy(), len(upload_df), 0

    # Convert PK columns to string for matching
    for col in pk:
        if col in master_df.columns:
            master_df[col] = master_df[col].astype(str)
        if col in upload_df.columns:
            upload_df[col] = upload_df[col].astype(str)

    # Find which upload rows match existing PKs
    upload_keys = upload_df[pk].apply(tuple, axis=1)
    master_keys = master_df[pk].apply(tuple, axis=1)

    existing_keys = set(master_keys)
    new_mask = ~upload_keys.isin(existing_keys)
    update_mask = upload_keys.isin(existing_keys)

    new_count = new_mask.sum()
    updated_count = update_mask.sum()

    # Remove old rows that will be updated
    if updated_count > 0:
        update_key_set = set(upload_keys[update_mask])
        keep_mask = ~master_keys.isin(update_key_set)
        master_df = master_df[keep_mask]

    # Append all upload rows (both new and updated)
    merged = pd.concat([master_df, upload_df], ignore_index=True)

    return merged, int(new_count), int(updated_count)
