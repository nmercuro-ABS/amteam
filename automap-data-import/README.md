# Automap Data Import

Internal tool for exporting TDS / BMC data to Excel with field-level documentation.

## Features
- SQL Server discovery via AbsWebSys
- Interactive login (Windows / macOS)
- Automated Excel exports with notes
- Dark / light mode UI

## Requirements
- Python 3.10+
- ODBC Driver 18 for SQL Server
- VPN access
- SQL permissions

## Run (dev)
```bash
source .venv/bin/activate
python src/automap_import.py
```

## Notes
- Generated Excel files are intentionally ignored by Git
- Authentication behavior differs between Windows and macOS
