# OAG Preliminary Information Request Tracker

Standalone tracker for the OAG "Preliminary Information Request" (FYE 30 June 2026).

## Run locally
    python -m pip install -r requirements.txt
    python -m streamlit run app.py

## Data
Everything (requests, attached documents, comments, activity log) is stored in ONE SQLite file,
`pbc_tracker.db`, created on first run. On first run the 97 requests from `pbc_seed.json` are loaded.
Set the `DATABASE_PATH` environment variable (e.g. a Railway volume such as `/data/pbc_tracker.db`)
so the data survives redeploys. Back up that single file to back up everything.

## Files
- `app.py` – entry point
- `pbc.py` – user interface
- `pbc_db.py` – database layer
- `pbc_parser.py` – Word import (keeps bold / lists / links) and HTML sanitiser
- `pbc_export.py` – Excel, Word and ZIP exports
- `pbc_seed.json` – the 97 requests parsed from the OAG Word document
