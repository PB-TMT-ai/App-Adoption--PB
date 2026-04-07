# App Adoption Tracking System

Track mobile app installations and adoption by dealers using Excel files.

## Quick Start

### 1. Install Python Dependencies

```bash
pip install -r requirements.txt
```

### 2. Create Templates

```bash
python scripts/run.py create-templates
```

This creates three Excel template files in the `templates/` folder:
- `dealer_info_template.xlsx` — Dealer information (name, location, contact)
- `install_events_template.xlsx` — App install/uninstall events
- `usage_metrics_template.xlsx` — Usage and engagement metrics

### 3. Fill In Your Data

1. Open a template from the `templates/` folder
2. Go to the **Data** sheet
3. Replace the sample data with your actual data
4. Save the file

**Tip:** Required columns are highlighted in green. Do not rename or reorder columns.

### 4. Upload Your Data

```bash
python scripts/run.py ingest path/to/your_file.xlsx
```

The system auto-detects the file type. You can also specify it:

```bash
python scripts/run.py ingest your_file.xlsx --type dealers
python scripts/run.py ingest your_file.xlsx --type install_events
python scripts/run.py ingest your_file.xlsx --type usage_metrics
```

**Upload order:** Upload dealers first, then install events, then usage metrics.

### 5. Generate Reports

```bash
python scripts/run.py report
```

This creates an Excel report in the `reports/` folder with these sheets:

| Sheet | What It Shows |
|-------|---------------|
| **Summary** | Overall adoption numbers (total dealers, adoption rate, active/inactive) |
| **Adoption Trend** | Monthly install/uninstall trend with adoption rate over time |
| **Dealer Detail** | One row per dealer with install status, last active date, engagement |
| **Engagement Breakdown** | Sessions, actions, and engagement scores per dealer |
| **At Risk** | Dealers who are installed but inactive or have low engagement |

## Commands

| Command | Description |
|---------|-------------|
| `python scripts/run.py create-templates` | Create Excel templates |
| `python scripts/run.py validate <file>` | Check a file for errors without importing |
| `python scripts/run.py ingest <file>` | Import data from an Excel file |
| `python scripts/run.py report` | Generate the adoption report |

## Folder Structure

```
templates/     Excel templates you fill in
uploads/       Your uploaded files (moved to processed/ after import)
data/          Master data files (CSV, managed by the system)
reports/       Generated report files
scripts/       Python scripts (do not modify)
```

## Troubleshooting

- **"Missing required columns"** — Make sure your file has all required columns (see template Instructions sheet)
- **"Could not auto-detect dataset type"** — Add `--type dealers`, `--type install_events`, or `--type usage_metrics`
- **Duplicate data** — Re-uploading the same file won't create duplicates; existing records are updated
