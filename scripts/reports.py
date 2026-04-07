"""
Generate adoption report Excel workbooks from master CSV data.
"""

import os
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from scripts.config import APP_ADOPTION_CSV, REPORT_DIR


def _load_data():
    """Load the master CSV file into a DataFrame."""
    if not os.path.exists(APP_ADOPTION_CSV) or os.path.getsize(APP_ADOPTION_CSV) == 0:
        return pd.DataFrame()

    df = pd.read_csv(APP_ADOPTION_CSV, dtype=str)

    # Parse numeric columns
    int_cols = ["# orders (total)", "# orders (via. app.)"]
    for col in int_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    numeric_cols = ["Order quantity (total)", "Quantity (via. app.)"]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].astype(str).str.strip().replace("-", "0"), errors="coerce").fillna(0)

    float_cols = ["% orders via. app.", "% order quantity (via. app.)"]
    for col in float_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


def _style_header(ws):
    """Apply header styling to the first row."""
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.freeze_panes = "A2"


def _auto_width(ws):
    """Auto-fit column widths."""
    for col_idx, col_cells in enumerate(ws.columns, 1):
        max_len = 0
        for cell in col_cells:
            try:
                val = str(cell.value) if cell.value is not None else ""
                max_len = max(max_len, len(val))
            except Exception:
                pass
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 3, 45)


def _add_metric_rows(ws, metrics):
    """Add label-value metric rows with bold labels."""
    bold = Font(bold=True, size=11)
    for label, value in metrics:
        ws.append([label, value])
        ws.cell(row=ws.max_row, column=1).font = bold


def _build_report_agg(df, group_col):
    """Build reporting aggregation grouped by group_col."""
    for col in ["# orders (total)", "# orders (via. app.)"]:
        if col not in df.columns:
            df[col] = 0
    for col in ["Order quantity (total)", "Quantity (via. app.)"]:
        if col not in df.columns:
            df[col] = 0.0

    agg = df.groupby(group_col).agg(
        target_dealers=("Account Sf ID", "count"),
        app_installs=("Status", lambda x: (x == "Installed").sum()),
        tech_issues=("Status", lambda x: (x == "Tech Issue").sum()),
        pending=("Status", lambda x: (x == "Not Installed").sum()),
        not_working=("Status", lambda x: (x == "Not Working").sum()),
        orders_total=("# orders (total)", "sum"),
        orders_app=("# orders (via. app.)", "sum"),
        qty_total=("Order quantity (total)", "sum"),
        qty_app=("Quantity (via. app.)", "sum"),
    ).reset_index()

    agg["orders_total"] = agg["orders_total"].fillna(0).astype(int)
    agg["orders_app"] = agg["orders_app"].fillna(0).astype(int)
    agg["qty_total"] = agg["qty_total"].fillna(0)
    agg["qty_app"] = agg["qty_app"].fillna(0)
    agg["pct_installs"] = (agg["app_installs"] / agg["target_dealers"].replace(0, pd.NA) * 100).fillna(0).round(0).astype(int)
    agg["pct_orders"] = (agg["orders_app"] / agg["orders_total"].replace(0, pd.NA) * 100).fillna(0).round(0).astype(int)
    agg["pct_qty"] = (agg["qty_app"] / agg["qty_total"].replace(0, pd.NA) * 100).fillna(0).round(0).astype(int)

    return agg


def generate_reports():
    """
    Generate the adoption report Excel workbook.

    Returns:
        Path to the generated report file.
    """
    df = _load_data()
    if df.empty:
        raise ValueError("No data found. Please ingest a data file first.")

    today = datetime.now().strftime("%Y-%m-%d")
    os.makedirs(REPORT_DIR, exist_ok=True)
    output_path = os.path.join(REPORT_DIR, f"adoption_report_{today}.xlsx")

    total = len(df)
    status_counts = df["Status"].value_counts()

    wb = Workbook()

    # ── Sheet 1: Summary ──
    ws = wb.active
    ws.title = "Summary"

    ws.append(["Metric", "Value"])

    installed = int(status_counts.get("Installed", 0))
    not_installed = int(status_counts.get("Not Installed", 0))
    uninstalled = int(status_counts.get("Uninstalled", 0))
    tech_issue = int(status_counts.get("Tech Issue", 0))
    not_working = int(status_counts.get("Not Working", 0))
    adoption_rate = round(installed / total * 100, 1) if total > 0 else 0
    ever_adopted = installed + uninstalled + tech_issue + not_working
    ever_adopted_rate = round(ever_adopted / total * 100, 1) if total > 0 else 0

    _add_metric_rows(ws, [
        ("Report Date", today),
        ("", ""),
        ("── Dealer Counts ──", ""),
        ("Total Dealers", total),
        ("Installed", installed),
        ("Not Installed", not_installed),
        ("Uninstalled", uninstalled),
        ("Tech Issue", tech_issue),
        ("Not Working", not_working),
        ("", ""),
        ("── Adoption Rates ──", ""),
        ("Currently Installed (%)", f"{adoption_rate}%"),
        ("Ever Adopted (%)", f"{ever_adopted_rate}%"),
        ("", ""),
        ("── Order Data ──", ""),
        ("Dealers with Orders via App",
         int((df["# orders (via. app.)"] > 0).sum()) if "# orders (via. app.)" in df.columns else "N/A"),
        ("Total Orders (all channels)",
         int(df["# orders (total)"].sum()) if "# orders (total)" in df.columns else "N/A"),
        ("Total Orders via App",
         int(df["# orders (via. app.)"].sum()) if "# orders (via. app.)" in df.columns else "N/A"),
    ])

    # Zone breakdown in summary
    ws.append(["", ""])
    ws.append(["── Zone Breakdown ──", ""])
    ws.cell(row=ws.max_row, column=1).font = Font(bold=True, size=11)
    ws.append(["Zone", "Total", "Installed", "Not Installed", "Uninstalled", "Tech Issue / Not Working", "Adoption Rate (%)"])
    ws.cell(row=ws.max_row, column=1).font = Font(bold=True)

    if "Zone" in df.columns:
        for zone in sorted(df["Zone"].dropna().unique()):
            z = df[df["Zone"] == zone]
            z_total = len(z)
            z_installed = int((z["Status"] == "Installed").sum())
            z_not_installed = int((z["Status"] == "Not Installed").sum())
            z_uninstalled = int((z["Status"] == "Uninstalled").sum())
            z_other = int(z["Status"].isin(["Tech Issue", "Not Working"]).sum())
            z_rate = round(z_installed / z_total * 100, 1) if z_total > 0 else 0
            ws.append([zone, z_total, z_installed, z_not_installed, z_uninstalled, z_other, f"{z_rate}%"])

    _style_header(ws)
    _auto_width(ws)

    # ── Sheet 2: Zone-wise Adoption ──
    ws_zone = wb.create_sheet("Zone-wise Adoption")
    ws_zone.append(["Zone", "State", "Total Dealers", "Installed", "Not Installed", "Uninstalled",
                     "Tech Issue", "Not Working", "Adoption Rate (%)"])

    if "Zone" in df.columns and "State" in df.columns:
        zone_state = df.groupby(["Zone", "State"]).agg(
            total=("Account Sf ID", "count"),
            installed=("Status", lambda x: (x == "Installed").sum()),
            not_installed=("Status", lambda x: (x == "Not Installed").sum()),
            uninstalled=("Status", lambda x: (x == "Uninstalled").sum()),
            tech_issue=("Status", lambda x: (x == "Tech Issue").sum()),
            not_working=("Status", lambda x: (x == "Not Working").sum()),
        ).reset_index()

        zone_state["rate"] = (zone_state["installed"] / zone_state["total"] * 100).round(1)
        zone_state = zone_state.sort_values(["Zone", "rate"], ascending=[True, False])

        for _, row in zone_state.iterrows():
            ws_zone.append([
                row["Zone"], row["State"], int(row["total"]),
                int(row["installed"]), int(row["not_installed"]), int(row["uninstalled"]),
                int(row["tech_issue"]), int(row["not_working"]), f"{row['rate']}%",
            ])

    _style_header(ws_zone)
    _auto_width(ws_zone)

    # ── Sheet 3: Distributor-level Report ──
    report_headers = [
        "Distributor", "State", "Target #", "Installs",
        "Tech Issue", "Pending", "N/W",
        "% Installs", "Orders (Total)", "Orders (App)",
        "% Orders", "Qty Total (MT)", "Qty App (MT)", "% Order Vol.",
    ]
    ws_dist = wb.create_sheet("Distributor-level Report")
    ws_dist.append(report_headers)

    pct_fill_green = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid")
    pct_fill_yellow = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    pct_fill_red = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")

    if "Distributor Name" in df.columns:
        dist_agg = _build_report_agg(df, "Distributor Name")
        if "State" in df.columns:
            smap = df.groupby("Distributor Name")["State"].first().reset_index()
            dist_agg = dist_agg.merge(smap, on="Distributor Name", how="left")
        else:
            dist_agg["State"] = ""
        dist_agg = dist_agg.sort_values("pct_installs", ascending=False)
        for _, row in dist_agg.iterrows():
            ws_dist.append([
                row["Distributor Name"], row.get("State", ""), int(row["target_dealers"]),
                int(row["app_installs"]), int(row["tech_issues"]), int(row["pending"]),
                int(row["not_working"]), f"{int(row['pct_installs'])}%",
                int(row["orders_total"]), int(row["orders_app"]), f"{int(row['pct_orders'])}%",
                round(row["qty_total"]), round(row["qty_app"]), f"{int(row['pct_qty'])}%",
            ])

    for ri in range(2, ws_dist.max_row + 1):
        for ci in [8, 11, 14]:
            cell = ws_dist.cell(row=ri, column=ci)
            try:
                val = float(str(cell.value).replace("%", ""))
            except (ValueError, TypeError):
                val = 0
            cell.fill = pct_fill_green if val >= 50 else (pct_fill_yellow if val >= 25 else pct_fill_red)
    _style_header(ws_dist)
    _auto_width(ws_dist)

    # ── Sheet 4: State-level Report ──
    state_report_headers = [
        "State", "Target #", "Installs",
        "Tech Issue", "Pending", "N/W",
        "% Installs", "Orders (Total)", "Orders (App)",
        "% Orders", "Qty Total (MT)", "Qty App (MT)", "% Order Vol.",
    ]
    ws_state = wb.create_sheet("State-level Report")
    ws_state.append(state_report_headers)

    if "State" in df.columns:
        st_agg = _build_report_agg(df, "State")
        st_agg = st_agg.sort_values("pct_installs", ascending=False)
        for _, row in st_agg.iterrows():
            ws_state.append([
                row["State"], int(row["target_dealers"]), int(row["app_installs"]),
                int(row["tech_issues"]), int(row["pending"]), int(row["not_working"]),
                f"{int(row['pct_installs'])}%", int(row["orders_total"]),
                int(row["orders_app"]), f"{int(row['pct_orders'])}%",
                round(row["qty_total"]), round(row["qty_app"]), f"{int(row['pct_qty'])}%",
            ])
        # Grand Total row
        t_d = int(st_agg["target_dealers"].sum())
        t_i = int(st_agg["app_installs"].sum())
        t_ti = int(st_agg["tech_issues"].sum())
        t_p = int(st_agg["pending"].sum())
        t_nw = int(st_agg["not_working"].sum())
        t_ot = int(st_agg["orders_total"].sum())
        t_oa = int(st_agg["orders_app"].sum())
        t_qt = round(st_agg["qty_total"].sum())
        t_qa = round(st_agg["qty_app"].sum())
        ws_state.append([
            "G total", t_d, t_i, t_ti, t_p, t_nw,
            f"{round(t_i / max(t_d, 1) * 100)}%", t_ot, t_oa,
            f"{round(t_oa / max(t_ot, 1) * 100)}%", t_qt, t_qa,
            f"{round(t_qa / max(t_qt, 1) * 100)}%",
        ])
        for ci in range(1, 14):
            ws_state.cell(row=ws_state.max_row, column=ci).font = Font(bold=True, size=11)

    for ri in range(2, ws_state.max_row + 1):
        for ci in [7, 10, 13]:
            cell = ws_state.cell(row=ri, column=ci)
            try:
                val = float(str(cell.value).replace("%", ""))
            except (ValueError, TypeError):
                val = 0
            cell.fill = pct_fill_green if val >= 50 else (pct_fill_yellow if val >= 25 else pct_fill_red)
    _style_header(ws_state)
    _auto_width(ws_state)

    # ── Sheet 4: Dealer Detail ──
    ws_detail = wb.create_sheet("Dealer Detail")

    detail_cols = [
        "Account Sf ID", "Name of the Dealer", "State", "Zone",
        "Distributor Name", "Account Owner As per SF", "Mobile No.",
        "# orders (total)", "# orders (via. app.)", "Status",
    ]
    available_cols = [c for c in detail_cols if c in df.columns]
    ws_detail.append(available_cols)

    # Sort: Installed first, then Uninstalled, Tech Issue, Not Working, Not Installed
    status_order = {"Installed": 0, "Uninstalled": 1, "Tech Issue": 2, "Not Working": 3, "Not Installed": 4}
    sorted_df = df.copy()
    sorted_df["_sort"] = sorted_df["Status"].map(status_order).fillna(5)
    sorted_df = sorted_df.sort_values(["_sort", "Name of the Dealer"])

    for _, row in sorted_df.iterrows():
        ws_detail.append([_cell_value(row.get(c)) for c in available_cols])

    _style_header(ws_detail)
    _auto_width(ws_detail)

    # ── Sheet 5: App Usage (Orders) ──
    ws_orders = wb.create_sheet("App Usage (Orders)")
    ws_orders.append([
        "Account Sf ID", "Name of the Dealer", "Zone", "Distributor Name",
        "# Orders (Total)", "# Orders (via App)", "% Orders via App",
        "Order Qty (Total)", "Qty (via App)", "% Qty via App", "Status",
    ])

    # Show dealers with any orders via app, sorted by orders descending
    if "# orders (via. app.)" in df.columns:
        app_users = df[df["# orders (via. app.)"] > 0].sort_values("# orders (via. app.)", ascending=False)
        for _, row in app_users.iterrows():
            ws_orders.append([
                _cell_value(row.get("Account Sf ID")),
                _cell_value(row.get("Name of the Dealer")),
                _cell_value(row.get("Zone")),
                _cell_value(row.get("Distributor Name")),
                _cell_value(row.get("# orders (total)")),
                _cell_value(row.get("# orders (via. app.)")),
                _pct_value(row.get("% orders via. app.")),
                _cell_value(row.get("Order quantity (total)")),
                _cell_value(row.get("Quantity (via. app.)")),
                _pct_value(row.get("% order quantity (via. app.)")),
                _cell_value(row.get("Status")),
            ])

        if app_users.empty:
            ws_orders.append(["No dealers have placed orders via the app yet."] + [""] * 10)

    _style_header(ws_orders)
    _auto_width(ws_orders)

    # ── Sheet 6: At Risk / Action Needed ──
    ws_risk = wb.create_sheet("At Risk - Action Needed")
    ws_risk.append([
        "Account Sf ID", "Name of the Dealer", "State", "Zone",
        "Distributor Name", "Account Owner As per SF", "Mobile No.",
        "Status", "Action Needed",
    ])

    risk_statuses = {"Uninstalled": "Re-engage: app was uninstalled",
                     "Tech Issue": "Resolve technical issue",
                     "Not Working": "Dealer not working"}

    risk_df = df[df["Status"].isin(risk_statuses.keys())].copy()
    risk_df["Action Needed"] = risk_df["Status"].map(risk_statuses)
    risk_df = risk_df.sort_values("Status")

    for _, row in risk_df.iterrows():
        ws_risk.append([
            _cell_value(row.get("Account Sf ID")),
            _cell_value(row.get("Name of the Dealer")),
            _cell_value(row.get("State")),
            _cell_value(row.get("Zone")),
            _cell_value(row.get("Distributor Name")),
            _cell_value(row.get("Account Owner As per SF")),
            _cell_value(row.get("Mobile No.")),
            _cell_value(row.get("Status")),
            _cell_value(row.get("Action Needed")),
        ])

    # Color-code by status
    status_colors = {
        "Not Working": PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid"),
        "Tech Issue": PatternFill(start_color="FFCC99", end_color="FFCC99", fill_type="solid"),
        "Uninstalled": PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid"),
    }
    status_col_idx = 8  # Column H = Status
    for row_idx in range(2, ws_risk.max_row + 1):
        status_val = ws_risk.cell(row=row_idx, column=status_col_idx).value
        if status_val in status_colors:
            for col_idx in range(1, 10):
                ws_risk.cell(row=row_idx, column=col_idx).fill = status_colors[status_val]

    _style_header(ws_risk)
    _auto_width(ws_risk)

    wb.save(output_path)
    return output_path


def _cell_value(val):
    """Convert a value for Excel output."""
    if pd.isna(val):
        return ""
    return val


def _pct_value(val):
    """Format a percentage value."""
    if pd.isna(val):
        return ""
    return f"{float(val):.1f}%"
