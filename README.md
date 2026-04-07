# App Adoption Tracking System

Track mobile app installations and adoption by dealers using Excel files.

## Quick Start

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Upload Your Data

Place your Excel file (e.g., `App adoption till 6th Apr.xlsx`) in the project folder, then run:

```bash
python scripts/run.py ingest "data/App adoption till 6th Apr.xlsx"
```

### 3. Generate Reports

```bash
python scripts/run.py report
```

The report is saved in the `reports/` folder.

## Expected Excel Format

Your Excel file should have these columns:

| Column | Required | Description |
|--------|----------|-------------|
| Account Sf ID | Yes | Unique dealer identifier |
| Name of the Dealer | Yes | Dealer name |
| State | No | State |
| Zone | No | North / Central / East |
| Distributor Name | No | Distributor |
| Account Owner As per SF | No | Account owner |
| GST No. | No | GST number |
| Mobile No. | No | Contact phone |
| Dealer classification | No | Active / Inactive |
| # orders (total) | No | Total orders |
| # orders (via. app.) | No | Orders placed through app |
| % orders via. app. | No | Percentage |
| Order quantity (total) | No | Total quantity ordered |
| Quantity (via. app.) | No | Quantity ordered via app |
| % order quantity (via. app.) | No | Percentage |
| Status | Yes | Installed / Not Installed / Uninstalled / Tech Issue / Not Working |

Use `-` for blank numeric values.

To create a blank template: `python scripts/run.py create-template`

## Report Sheets

| Sheet | What It Shows |
|-------|---------------|
| **Summary** | Total dealers, status counts, adoption rate, zone breakdown |
| **Zone-wise Adoption** | Adoption rate by zone and state |
| **Distributor-wise Adoption** | Which distributors have best/worst adoption |
| **Dealer Detail** | All dealers sorted by status |
| **App Usage (Orders)** | Dealers placing orders through the app |
| **At Risk - Action Needed** | Uninstalled + Tech Issue + Not Working dealers for follow-up |

## Commands

| Command | Description |
|---------|-------------|
| `python scripts/run.py ingest <file>` | Import data from Excel |
| `python scripts/run.py report` | Generate adoption report |
| `python scripts/run.py validate <file>` | Check a file for errors |
| `python scripts/run.py create-template` | Create a blank Excel template |
