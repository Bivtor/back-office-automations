# Lockbox Parquet Refresh

Power BI data preparation for the lockbox management reports. Converts the "Manager Report Detail" CSV exports into Parquet so the reports load faster and use far less memory, ports the Power Query and DAX transformations into pandas so they run once at refresh time instead of on every report open, and recomputes the current month's row of the KPI benchmarks workbook.

```
refresh.bat             double-click this; finds Python, installs pandas/pyarrow/openpyxl if missing, runs the script
lockbox_to_parquet.py   the converter
```

## What it does

- **Three CSV profiles**, each with its own column pruning and value cleanup ported from the Power Query that used to read it: the 12-month folder (monthly report), the current raw data folder (weekly report), and the full history since 2020. One `.parquet` per source `.csv`, rebuilt only when the CSV is newer or the derived-column logic version changed, so a monthly rotation only costs the new file.
- **Derived columns ported from DAX**: hold durations, hold reason categories, business-analyst user groupings, earliest-remittance timestamps per batch, and the hour-fraction gaps the processing-rate measures depend on. Two long-standing DAX bugs are fixed and documented at the top of the script (blank-date arithmetic that produced million-hour gaps, and a dead null check that mislabelled holds); `BUG_COMPAT_MODE` reproduces the old output for diffing.
- **Control sheet data**: rebuilds one Parquet file from the daily bank control-sheet workbooks produced by [`daily-control-sheet`](../daily-control-sheet/).
- **KPI benchmarks**: recomputes the trailing-13-month row in the benchmarks workbook (processing and auditing rates in batches per hour) and prints old versus new values, with a calibration mode that fits the script's figures to the historical rows before writing anything.
- `--audit` compares CSV schemas across the folders and exits, which is how a new export layout gets caught before it silently changes a report.

## Usage

```
python lockbox_to_parquet.py                 # convert everything, then benchmarks
python lockbox_to_parquet.py --audit         # compare CSV schemas, then exit
python lockbox_to_parquet.py --force         # rebuild regardless of timestamps
python lockbox_to_parquet.py --only weekly   # one profile (monthly / weekly / alldata / controlsheet)
python lockbox_to_parquet.py --benchmarks-only --dry-run   # calibrate, write nothing
python lockbox_to_parquet.py --no-benchmarks                # Parquet only
```

Paths at the top of the script point at the shared-drive folders the reports read from. Employer names in the value-fix tables are placeholders here (`Sample ISD`, `Sample Health`); the originals corrected known misspellings and merged plan variants of the same employer.
