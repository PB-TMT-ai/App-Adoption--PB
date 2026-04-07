#!/usr/bin/env python3
"""
CLI entry point for the App Adoption Tracking System.

Usage:
    python scripts/run.py create-templates
    python scripts/run.py validate <file> [--type TYPE]
    python scripts/run.py ingest <file> [--type TYPE]
    python scripts/run.py report
"""

import argparse
import os
import sys

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation

# Add project root to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from scripts.config import (
    SCHEMAS, SAMPLE_DATA, TEMPLATE_DIR,
    get_column_names, detect_dataset_type,
)
from scripts.validate import validate_upload
from scripts.ingest import ingest_file
from scripts.reports import generate_reports


def cmd_create_templates(_args):
    """Create Excel template files with headers, sample data, and instructions."""
    os.makedirs(TEMPLATE_DIR, exist_ok=True)

    for dtype, schema in SCHEMAS.items():
        template_path = os.path.join(TEMPLATE_DIR, schema["template_name"])
        wb = Workbook()

        # Data sheet
        ws = wb.active
        ws.title = "Data"

        columns = get_column_names(dtype)
        required = {col[0] for col in schema["columns"] if col[2]}

        # Header row
        header_font = Font(bold=True, color="FFFFFF", size=11)
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        req_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

        for col_idx, col_name in enumerate(columns, 1):
            cell = ws.cell(row=1, column=col_idx, value=col_name)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")

        # Sample data rows
        samples = SAMPLE_DATA.get(dtype, [])
        for row_idx, sample in enumerate(samples, 2):
            for col_idx, col_name in enumerate(columns, 1):
                value = sample.get(col_name, "")
                ws.cell(row=row_idx, column=col_idx, value=value)

        # Highlight required columns in row 2+
        for col_idx, col_name in enumerate(columns, 1):
            if col_name in required:
                for row_idx in range(2, len(samples) + 2):
                    ws.cell(row=row_idx, column=col_idx).fill = req_fill

        # Auto-fit column widths
        for col_idx, col_name in enumerate(columns, 1):
            ws.column_dimensions[chr(64 + col_idx) if col_idx <= 26 else f"A{chr(64 + col_idx - 26)}"].width = max(len(col_name) + 4, 15)

        ws.freeze_panes = "A2"

        # Add data validation for event_type in install_events
        if dtype == "install_events":
            event_type_col = columns.index("event_type") + 1
            dv = DataValidation(type="list", formula1='"install,uninstall"', allow_blank=False)
            dv.error = "Please enter 'install' or 'uninstall'"
            dv.errorTitle = "Invalid event type"
            ws.add_data_validation(dv)
            for row_idx in range(2, 1002):  # Apply to first 1000 rows
                dv.add(ws.cell(row=row_idx, column=event_type_col))

        # Instructions sheet
        ws_instr = wb.create_sheet("Instructions")
        instructions = _get_instructions(dtype, columns, required)
        for row_idx, line in enumerate(instructions, 1):
            cell = ws_instr.cell(row=row_idx, column=1, value=line)
            if row_idx == 1:
                cell.font = Font(bold=True, size=14)
            elif line.startswith("*"):
                cell.font = Font(bold=True, size=11)
        ws_instr.column_dimensions["A"].width = 80

        wb.save(template_path)
        print(f"  Created: {template_path}")

    print(f"\nAll templates created in: {TEMPLATE_DIR}")
    print("Fill in the 'Data' sheet and save. Then run: python scripts/run.py ingest <file>")


def _get_instructions(dtype, columns, required):
    """Generate instruction text for a template."""
    type_labels = {
        "dealers": "Dealer Information",
        "install_events": "App Install/Uninstall Events",
        "usage_metrics": "Usage & Engagement Metrics",
    }
    lines = [
        f"Template: {type_labels.get(dtype, dtype)}",
        "",
        "* How to use this template:",
        "1. Go to the 'Data' sheet",
        "2. Replace the sample data with your actual data (one row per record)",
        "3. Do NOT rename or reorder the columns",
        "4. Save the file",
        f"5. Run: python scripts/run.py ingest <this_file.xlsx> --type {dtype}",
        "",
        f"* Required columns (highlighted in green):",
    ]
    for col in columns:
        marker = " [REQUIRED]" if col in required else ""
        lines.append(f"  - {col}{marker}")

    if dtype == "install_events":
        lines += [
            "",
            "* event_type must be either 'install' or 'uninstall'",
            "* event_date should be in YYYY-MM-DD format (e.g., 2025-03-15)",
        ]
    elif dtype == "usage_metrics":
        lines += [
            "",
            "* period_start and period_end define the measurement window",
            "* engagement_score should be between 0 and 100",
            "* features_used: comma-separated list (e.g., 'inventory,messaging')",
        ]
    elif dtype == "dealers":
        lines += [
            "",
            "* dealer_id must be unique for each dealer",
            "* onboarded_date should be in YYYY-MM-DD format",
        ]

    return lines


def cmd_validate(args):
    """Validate an uploaded Excel file."""
    file_path = args.file
    dataset_type = args.type

    if not os.path.exists(file_path):
        print(f"ERROR: File not found: {file_path}")
        sys.exit(1)

    # Auto-detect type if not specified
    if not dataset_type:
        try:
            df = pd.read_excel(file_path, engine="openpyxl", nrows=0)
            dataset_type = detect_dataset_type(df.columns)
            if dataset_type:
                print(f"Auto-detected dataset type: {dataset_type}")
            else:
                print("ERROR: Could not auto-detect dataset type. Please specify with --type")
                sys.exit(1)
        except Exception as e:
            print(f"ERROR: Could not read file: {e}")
            sys.exit(1)

    is_valid, messages = validate_upload(file_path, dataset_type)

    print(f"\nValidation results for: {os.path.basename(file_path)}")
    print(f"Dataset type: {dataset_type}")
    print("-" * 50)
    for msg in messages:
        print(f"  {msg}")
    print("-" * 50)

    if is_valid:
        print("RESULT: PASSED - File is ready for ingestion.")
        print(f"Next step: python scripts/run.py ingest {file_path} --type {dataset_type}")
    else:
        print("RESULT: FAILED - Please fix the errors above and try again.")
        sys.exit(1)


def cmd_ingest(args):
    """Ingest an uploaded Excel file into master data."""
    file_path = args.file
    dataset_type = args.type

    if not os.path.exists(file_path):
        print(f"ERROR: File not found: {file_path}")
        sys.exit(1)

    # Auto-detect type if not specified
    if not dataset_type:
        try:
            df = pd.read_excel(file_path, engine="openpyxl", nrows=0)
            dataset_type = detect_dataset_type(df.columns)
            if dataset_type:
                print(f"Auto-detected dataset type: {dataset_type}")
            else:
                print("ERROR: Could not auto-detect dataset type. Please specify with --type")
                sys.exit(1)
        except Exception as e:
            print(f"ERROR: Could not read file: {e}")
            sys.exit(1)

    print(f"Ingesting: {os.path.basename(file_path)} (type: {dataset_type})")
    print("-" * 50)

    try:
        rows_affected, messages = ingest_file(file_path, dataset_type)
        for msg in messages:
            print(f"  {msg}")
        print("-" * 50)
        print(f"SUCCESS: {rows_affected} row(s) processed.")
    except ValueError as e:
        print(f"\n{e}")
        sys.exit(1)


def cmd_report(_args):
    """Generate adoption reports."""
    print("Generating adoption report...")
    print("-" * 50)

    try:
        output_path = generate_reports()
        print(f"  Report saved to: {output_path}")
        print("-" * 50)
        print("SUCCESS: Open the Excel file to view the report.")
    except Exception as e:
        print(f"ERROR: Failed to generate report: {e}")
        sys.exit(1)


def main():
    parser = argparse.ArgumentParser(
        description="App Adoption Tracking System",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python scripts/run.py create-templates
  python scripts/run.py validate uploads/dealers.xlsx --type dealers
  python scripts/run.py ingest uploads/dealers.xlsx --type dealers
  python scripts/run.py report
        """,
    )
    subparsers = parser.add_subparsers(dest="command", help="Command to run")

    # create-templates
    sub_templates = subparsers.add_parser("create-templates", help="Create Excel template files")
    sub_templates.set_defaults(func=cmd_create_templates)

    # validate
    sub_validate = subparsers.add_parser("validate", help="Validate an uploaded Excel file")
    sub_validate.add_argument("file", help="Path to the Excel file")
    sub_validate.add_argument("--type", choices=list(SCHEMAS.keys()),
                              help="Dataset type (auto-detected if omitted)")
    sub_validate.set_defaults(func=cmd_validate)

    # ingest
    sub_ingest = subparsers.add_parser("ingest", help="Ingest an uploaded Excel file")
    sub_ingest.add_argument("file", help="Path to the Excel file")
    sub_ingest.add_argument("--type", choices=list(SCHEMAS.keys()),
                            help="Dataset type (auto-detected if omitted)")
    sub_ingest.set_defaults(func=cmd_ingest)

    # report
    sub_report = subparsers.add_parser("report", help="Generate adoption report")
    sub_report.set_defaults(func=cmd_report)

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(1)

    args.func(args)


if __name__ == "__main__":
    main()
