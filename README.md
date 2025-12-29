# Kowsar Stock Agent (UI Automation → Clean Excel → SQLite → Search API)

## Problem
Kowsar accounting software did not provide an API for inventory access, so this project automates the UI to export inventory reports,
then cleans/normalizes Persian text and converts the output into a SQLite database for downstream AI/search usage.

## What it does
- Logs into Kowsar and navigates to the inventory report using image-based UI automation (PyAutoGUI).
- Exports an Excel report on an interval (default: 15 minutes).
- Normalizes Persian/Arabic characters and cleans fields for consistent search.
- Converts the cleaned Excel into SQLite (`products` table).
- Exposes a FastAPI `/search` endpoint for the chatbot brain to query the inventory DB.

## Components
- `src/agent.py`: UI automation + export + normalization + DB build
- `src/dbmaker.py`: Excel → SQLite utilities (and column cleaning)
- `src/stock_api.py`: FastAPI search service over SQLite DB
- `src/build_brand_aliases.py`: builds brand aliases for better search matching

## Setup (Windows)
1) Install Python 3.10+
2) Create venv and install:
```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
