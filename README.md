# Excel Google Sheets Sync

A Python CLI for bidirectional synchronization between Excel files and Google Sheets. It formats exported workbooks and creates a backup sheet before imports.

## Features

- Export Google Sheets data to `.xlsx`.
- Import `.xlsx` data into Google Sheets with structure validation.
- Create a backup sheet before import.
- Format headers, borders, alignment, and column widths.
- Apply conditional highlighting to numeric values.

## Stack

- Python 3.10+
- gspread and Google Auth
- openpyxl
- python-dotenv

## Quick start

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
cp creds.json.example creds.json
```

Set the spreadsheet ID and sheet name in `.env`, then replace `creds.json` with a Google Service Account key. Share the spreadsheet with the `client_email` address from that key.

```bash
python sync.py --export
python sync.py --import --file report.xlsx
```

## Project structure

```text
sync.py             CLI and synchronization entry point
sheets_client.py    Google Sheets API operations
excel_handler.py    Excel reading and writing
formatter.py        workbook formatting rules
config.py           configuration loading
requirements.txt    Python dependencies
```

## License

MIT
