"""
Generate adoption report Excel workbooks from master CSV data.
"""

import os
from datetime import datetime, timedelta

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, numbers
from openpyxl.utils import get_column_letter

from scripts.config import (
    DEALERS_CSV, INSTALL_EVENTS_CSV, USAGE_METRICS_CSV,
    REPORT_DIR, INACTIVE_DAYS, AT_RISK_DAYS,
    ENGAGEMENT_HIGH, ENGAGEMENT_MEDIUM,
)


def _load_data():
    """Load all master CSV files into DataFrames."""
    dealers = pd.DataFrame()
    events = pd.DataFrame()
    usage = pd.DataFrame()

    if os.path.exists(DEALERS_CSV) and os.path.getsize(DEALERS_CSV) > 0:
        dealers = pd.read_csv(DEALERS_CSV, dtype=str)
        if "onboarded_date" in dealers.columns:
            dealers["onboarded_date"] = pd.to_datetime(dealers["onboarded_date"], errors="coerce")

    if os.path.exists(INSTALL_EVENTS_CSV) and os.path.getsize(INSTALL_EVENTS_CSV) > 0:
        events = pd.read_csv(INSTALL_EVENTS_CSV, dtype=str)
        events["event_date"] = pd.to_datetime(events["event_date"], errors="coerce")

    if os.path.exists(USAGE_METRICS_CSV) and os.path.getsize(USAGE_METRICS_CSV) > 0:
        usage = pd.read_csv(USAGE_METRICS_CSV)
        for col in ["period_start", "period_end", "last_active_date"]:
            if col in usage.columns:
                usage[col] = pd.to_datetime(usage[col], errors="coerce")
        for col in ["login_count", "sessions_count", "actions_performed"]:
            if col in usage.columns:
                usage[col] = pd.to_numeric(usage[col], errors="coerce").fillna(0).astype(int)
        for col in ["avg_session_minutes", "engagement_score"]:
            if col in usage.columns:
                usage[col] = pd.to_numeric(usage[col], errors="coerce")

    return dealers, events, usage


def _style_header(ws):
    """Apply header styling to the first row of a worksheet."""
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")

    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    ws.freeze_panes = "A2"


def _auto_width(ws):
    """Auto-fit column widths based on content."""
    for col_idx, col_cells in enumerate(ws.columns, 1):
        max_len = 0
        for cell in col_cells:
            try:
                val = str(cell.value) if cell.value is not None else ""
                max_len = max(max_len, len(val))
            except Exception:
                pass
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 3, 40)


def _get_dealer_install_status(events):
    """
    Determine current install status for each dealer based on events.

    Returns DataFrame with: dealer_id, current_status, install_date, app_version
    """
    if events.empty:
        return pd.DataFrame(columns=["dealer_id", "current_status", "install_date", "app_version"])

    # Sort by date, take the latest event per dealer
    sorted_events = events.sort_values("event_date")
    latest = sorted_events.groupby("dealer_id").last().reset_index()

    result = latest[["dealer_id"]].copy()
    result["current_status"] = latest["event_type"].apply(
        lambda x: "installed" if str(x).lower() == "install" else "uninstalled"
    )
    result["install_date"] = latest["event_date"]
    result["app_version"] = latest.get("app_version", "")

    return result


def _engagement_tier(score):
    """Return engagement tier label based on score."""
    if pd.isna(score):
        return "N/A"
    if score >= ENGAGEMENT_HIGH:
        return "High"
    if score >= ENGAGEMENT_MEDIUM:
        return "Medium"
    return "Low"


def generate_reports():
    """
    Generate the adoption report Excel workbook.

    Returns:
        Path to the generated report file.
    """
    dealers, events, usage = _load_data()
    today = datetime.now().date()

    os.makedirs(REPORT_DIR, exist_ok=True)
    output_path = os.path.join(REPORT_DIR, f"adoption_report_{today.isoformat()}.xlsx")

    wb = Workbook()

    # ── Sheet 1: Summary ──
    ws_summary = wb.active
    ws_summary.title = "Summary"

    total_dealers = len(dealers)
    install_status = _get_dealer_install_status(events)
    adopted = len(install_status[install_status["current_status"] == "installed"]) if not install_status.empty else 0
    uninstalled = len(install_status[install_status["current_status"] == "uninstalled"]) if not install_status.empty else 0
    ever_installed = len(install_status) if not install_status.empty else 0
    adoption_rate = (ever_installed / total_dealers * 100) if total_dealers > 0 else 0

    # Active/inactive from usage data
    active_count = 0
    inactive_count = 0
    if not usage.empty and "last_active_date" in usage.columns:
        latest_usage = usage.sort_values("period_end").groupby("dealer_id").last().reset_index()
        cutoff = pd.Timestamp(today - timedelta(days=INACTIVE_DAYS))
        active_ids = latest_usage[latest_usage["last_active_date"] >= cutoff]["dealer_id"]
        active_count = len(active_ids)
        # Inactive = installed but not active
        installed_ids = set(install_status[install_status["current_status"] == "installed"]["dealer_id"])
        inactive_count = len(installed_ids - set(active_ids))

    metrics = [
        ("Metric", "Value"),
        ("Report Date", str(today)),
        ("Total Dealers", total_dealers),
        ("Ever Installed", ever_installed),
        ("Currently Installed", adopted),
        ("Uninstalled (Churned)", uninstalled),
        ("Never Installed", total_dealers - ever_installed),
        ("Adoption Rate (%)", round(adoption_rate, 1)),
        ("Active (last 30 days)", active_count),
        ("Inactive (installed, no activity 30+ days)", inactive_count),
    ]
    for row in metrics:
        ws_summary.append(row)

    _style_header(ws_summary)
    _auto_width(ws_summary)

    # ── Sheet 2: Adoption Trend ──
    ws_trend = wb.create_sheet("Adoption Trend")

    if not events.empty:
        events_copy = events.copy()
        events_copy["month"] = events_copy["event_date"].dt.to_period("M")
        monthly = events_copy.groupby(["month", "event_type"]).size().unstack(fill_value=0).reset_index()
        monthly.columns = [str(c) for c in monthly.columns]

        if "install" not in monthly.columns:
            monthly["install"] = 0
        if "uninstall" not in monthly.columns:
            monthly["uninstall"] = 0

        monthly = monthly.sort_values("month")
        monthly["cumulative_installs"] = monthly["install"].cumsum()
        monthly["cumulative_uninstalls"] = monthly["uninstall"].cumsum()
        monthly["net_installed"] = monthly["cumulative_installs"] - monthly["cumulative_uninstalls"]
        monthly["adoption_rate_pct"] = (
            (monthly["net_installed"] / total_dealers * 100).round(1) if total_dealers > 0 else 0
        )

        ws_trend.append(["Month", "New Installs", "Uninstalls", "Cumulative Installs", "Net Installed", "Adoption Rate (%)"])
        for _, row in monthly.iterrows():
            ws_trend.append([
                str(row["month"]),
                int(row["install"]),
                int(row["uninstall"]),
                int(row["cumulative_installs"]),
                int(row["net_installed"]),
                float(row["adoption_rate_pct"]),
            ])
    else:
        ws_trend.append(["Month", "New Installs", "Uninstalls", "Cumulative Installs", "Net Installed", "Adoption Rate (%)"])
        ws_trend.append(["No data available", "", "", "", "", ""])

    _style_header(ws_trend)
    _auto_width(ws_trend)

    # ── Sheet 3: Dealer Detail ──
    ws_detail = wb.create_sheet("Dealer Detail")

    detail_headers = [
        "Dealer ID", "Dealer Name", "Dealer Group", "Region",
        "Install Status", "Install Date", "App Version",
        "Last Active Date", "Total Logins", "Avg Engagement Score",
        "Days Since Last Activity", "Status Label"
    ]
    ws_detail.append(detail_headers)

    if not dealers.empty:
        detail_df = dealers[["dealer_id", "dealer_name"]].copy()
        if "dealer_group" in dealers.columns:
            detail_df["dealer_group"] = dealers["dealer_group"]
        else:
            detail_df["dealer_group"] = ""
        if "region" in dealers.columns:
            detail_df["region"] = dealers["region"]
        else:
            detail_df["region"] = ""

        # Merge install status
        if not install_status.empty:
            detail_df = detail_df.merge(
                install_status[["dealer_id", "current_status", "install_date", "app_version"]],
                on="dealer_id", how="left"
            )
        else:
            detail_df["current_status"] = "not installed"
            detail_df["install_date"] = None
            detail_df["app_version"] = ""

        detail_df["current_status"] = detail_df["current_status"].fillna("not installed")

        # Merge usage data
        if not usage.empty:
            latest_usage = usage.sort_values("period_end").groupby("dealer_id").last().reset_index()
            agg_usage = usage.groupby("dealer_id").agg(
                total_logins=("login_count", "sum"),
                avg_engagement=("engagement_score", "mean"),
            ).reset_index()

            detail_df = detail_df.merge(
                latest_usage[["dealer_id", "last_active_date"]],
                on="dealer_id", how="left"
            )
            detail_df = detail_df.merge(agg_usage, on="dealer_id", how="left")
        else:
            detail_df["last_active_date"] = None
            detail_df["total_logins"] = 0
            detail_df["avg_engagement"] = None

        # Compute days since last activity
        detail_df["days_since"] = detail_df["last_active_date"].apply(
            lambda x: (pd.Timestamp(today) - x).days if pd.notna(x) else None
        )

        # Status label
        def status_label(row):
            if row["current_status"] == "uninstalled":
                return "Churned"
            if row["current_status"] == "not installed":
                return "Not Installed"
            if pd.isna(row.get("days_since")) or row.get("days_since") is None:
                return "No Usage Data"
            if row["days_since"] > INACTIVE_DAYS:
                return "Inactive"
            return "Active"

        detail_df["status_label"] = detail_df.apply(status_label, axis=1)

        # Sort: Active first, then Inactive, Churned, Not Installed
        status_order = {"Active": 0, "Inactive": 1, "No Usage Data": 2, "Churned": 3, "Not Installed": 4}
        detail_df["sort_key"] = detail_df["status_label"].map(status_order)
        detail_df = detail_df.sort_values(["sort_key", "dealer_name"]).drop(columns=["sort_key"])

        for _, row in detail_df.iterrows():
            ws_detail.append([
                str(row.get("dealer_id", "")),
                str(row.get("dealer_name", "")),
                str(row.get("dealer_group", "")),
                str(row.get("region", "")),
                str(row.get("current_status", "")),
                row.get("install_date", "").strftime("%Y-%m-%d") if pd.notna(row.get("install_date")) else "",
                str(row.get("app_version", "")),
                row.get("last_active_date", "").strftime("%Y-%m-%d") if pd.notna(row.get("last_active_date")) else "",
                int(row.get("total_logins", 0)) if pd.notna(row.get("total_logins")) else 0,
                round(float(row.get("avg_engagement", 0)), 1) if pd.notna(row.get("avg_engagement")) else "",
                int(row.get("days_since", 0)) if pd.notna(row.get("days_since")) else "",
                str(row.get("status_label", "")),
            ])

    _style_header(ws_detail)
    _auto_width(ws_detail)

    # ── Sheet 4: Engagement Breakdown ──
    ws_engage = wb.create_sheet("Engagement Breakdown")

    engage_headers = [
        "Dealer ID", "Dealer Name", "Total Sessions", "Total Actions",
        "Avg Session (min)", "Engagement Score", "Engagement Tier"
    ]
    ws_engage.append(engage_headers)

    if not usage.empty and not dealers.empty:
        engage_df = usage.groupby("dealer_id").agg(
            total_sessions=("sessions_count", "sum"),
            total_actions=("actions_performed", "sum"),
            avg_session_min=("avg_session_minutes", "mean"),
            avg_engagement=("engagement_score", "mean"),
        ).reset_index()

        engage_df = engage_df.merge(
            dealers[["dealer_id", "dealer_name"]], on="dealer_id", how="left"
        )
        engage_df["tier"] = engage_df["avg_engagement"].apply(_engagement_tier)
        engage_df = engage_df.sort_values("avg_engagement", ascending=False)

        for _, row in engage_df.iterrows():
            ws_engage.append([
                str(row["dealer_id"]),
                str(row.get("dealer_name", "")),
                int(row["total_sessions"]),
                int(row["total_actions"]),
                round(float(row["avg_session_min"]), 1) if pd.notna(row["avg_session_min"]) else "",
                round(float(row["avg_engagement"]), 1) if pd.notna(row["avg_engagement"]) else "",
                str(row["tier"]),
            ])

    _style_header(ws_engage)
    _auto_width(ws_engage)

    # ── Sheet 5: At Risk ──
    ws_risk = wb.create_sheet("At Risk")

    risk_headers = [
        "Dealer ID", "Dealer Name", "Region", "Last Active Date",
        "Days Since Activity", "Engagement Score", "Risk Reason"
    ]
    ws_risk.append(risk_headers)

    if not dealers.empty and not install_status.empty:
        installed_ids = set(install_status[install_status["current_status"] == "installed"]["dealer_id"])
        at_risk_dealers = dealers[dealers["dealer_id"].isin(installed_ids)].copy()

        if not usage.empty:
            latest_usage = usage.sort_values("period_end").groupby("dealer_id").last().reset_index()
            at_risk_dealers = at_risk_dealers.merge(
                latest_usage[["dealer_id", "last_active_date", "engagement_score"]],
                on="dealer_id", how="left"
            )
        else:
            at_risk_dealers["last_active_date"] = None
            at_risk_dealers["engagement_score"] = None

        at_risk_dealers["days_since"] = at_risk_dealers["last_active_date"].apply(
            lambda x: (pd.Timestamp(today) - x).days if pd.notna(x) else None
        )

        # Filter: no activity in AT_RISK_DAYS+ or low engagement
        risk_rows = at_risk_dealers[
            (at_risk_dealers["days_since"].isna()) |
            (at_risk_dealers["days_since"] > AT_RISK_DAYS) |
            (at_risk_dealers["engagement_score"] < ENGAGEMENT_MEDIUM)
        ].copy()

        def risk_reason(row):
            reasons = []
            if pd.isna(row.get("days_since")):
                reasons.append("No usage data")
            elif row["days_since"] > AT_RISK_DAYS:
                reasons.append(f"Inactive {int(row['days_since'])} days")
            if pd.notna(row.get("engagement_score")) and row["engagement_score"] < ENGAGEMENT_MEDIUM:
                reasons.append(f"Low engagement ({row['engagement_score']:.0f})")
            return "; ".join(reasons) if reasons else "At risk"

        risk_rows["risk_reason"] = risk_rows.apply(risk_reason, axis=1)
        risk_rows = risk_rows.sort_values("days_since", ascending=False, na_position="first")

        for _, row in risk_rows.iterrows():
            ws_risk.append([
                str(row["dealer_id"]),
                str(row.get("dealer_name", "")),
                str(row.get("region", "")),
                row["last_active_date"].strftime("%Y-%m-%d") if pd.notna(row.get("last_active_date")) else "Never",
                int(row["days_since"]) if pd.notna(row.get("days_since")) else "N/A",
                round(float(row["engagement_score"]), 1) if pd.notna(row.get("engagement_score")) else "N/A",
                str(row["risk_reason"]),
            ])

    _style_header(ws_risk)
    _auto_width(ws_risk)

    # Save workbook
    wb.save(output_path)
    return output_path
