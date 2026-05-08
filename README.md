# Disney Airtable → Excel Calendar Automation

## Overview

This project automates the export of Disney YouTube planning data from Airtable into Excel calendar templates.

Airtable is the source of truth.  
Excel is the provider-facing output layer.

Current supported outputs:

- `FULL CALENDAR YT DC EMEA`
- `MONTHLY OVERVIEW YT DC EMEA`

---

## Project Structure

```text
templates/   Excel templates (.xltx)
data/        Sample or exported Airtable JSON
output/      Generated Excel files
src/         Python source code
.env         Local credentials and settings, not committed

```
---
## Setup
1. Create virtual environment

Windows:
```bash
python -m venv .venv
.venv\Scripts\activate
```
Mac/Linux:
```bash
python3 -m venv .venv
source .venv/bin/activate
```
---
2. Install dependencies
```bash
pip install -r requirements.txt
```
---
3. Configure environment

Create .env based on .env.example.

Required variables:
AIRTABLE_API_TOKEN=your_token_here
AIRTABLE_BASE_ID=your_base_id_here
AIRTABLE_TABLE_ID=your_activities_table_id_here
AIRTABLE_VIEW=
ENABLE_MEDIAPLANNING=false

---
## Airtable Field Requirements

Main table: Activities

Important fields:
Brand
Property
Activity
Content Type
Live Date
YouTube Title
YouTube Description
Notes
TX ID / LINK
TC IN
TC OUT
Strategy
Region
Excluded Markets
YouTube Channel
GeoFenced?
Owner Hub
Priority
Month
Week Number (Year)
Deadline Delivery

Current expected Month format: ``YYYY-MM``
Example: ``2026-06``

Current expected Week Number (Year) format: ``YYYY-WXX``
Example: ``026-W23``

---

## Run Export
Disney Channel, June 2026
```bash
.venv\Scripts\python.exe src\populate_calendar.py --template "templates\X_Social Media Calendar DC YouTube EMEA - month YYYY.xltx" --output "output\DC_EMEA_JUN26.xlsx" --channel "Disney Channel" --month "2026-06" --log-level INFO
```

or without template (default with Channel/Brand):
```bash
.venv\Scripts\python.exe src\populate_calendar.py --output "output\test_widths_july.xlsx" --channel "Disney Channel" --month "2026-07" --log-level INFO
```

With debug logs:
```bash
.venv\Scripts\python.exe src\populate_calendar.py --template "templates\X_Social Media Calendar DC YouTube EMEA - month YYYY.xltx" --output "output\debug_output.xlsx" --channel "Disney Channel" --month "2026-06" --log-level DEBUG
```

Using local JSON:
```bash
.venv\Scripts\python.exe src\populate_calendar.py --template "templates\X_Social Media Calendar DC YouTube EMEA - month YYYY.xltx" --output "output\test_exported.xlsx" --json "data\exported_airtable.json" --channel "Disney Channel" --month "2026-06" --log-level DEBUG
```
---

## Streamlit App

Run the local interface:
```bash
streamlit run src\ui\streamlit_app.py
```
The app is intended to let users select:

- Template
- Output file name
- Brand / Channel
- Month
- Owner hub options
- Whether to include Monthly Overview
---

## Current Notes
- Brand is now the Airtable field used to separate Disney Channel and Disney Jr.
- The script still uses the CLI argument --channel for compatibility, but internally this should be refactored to --brand.
- Excluded Markets replaces the previous Notes-based exception logic.
- Excluded markets are appended automatically to the Excel Comments section as Except France, Except MENA, etc.
- Mediaplanning and Creative Services can be enabled through .env.