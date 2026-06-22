# Disney Airtable Calendar Exporter

## Overview

This project exports YouTube planning data from Airtable into Disney Excel calendar templates used across EMEA markets.

The tool currently supports:

* Disney Channel (DC)
* Disney Jr. (DJR)

The project generates:

1. Full Calendar (detailed publishing calendar)
2. Monthly Overview (high-level planning view)
3. Country calendars (via rebalance automation)

The goal is to reduce manual Excel work while preserving the original template formatting, formulas, merged cells, colors, and layout.

---

# Current Status

## Disney Channel

Status: Stable

Features:

* Full Calendar export
* Monthly Overview export
* Country rollout logic
* Mediaplanning support
* QA checks
* Streamlit UI support

## Disney Jr.

Status: Functional but still evolving

Features:

* Full Calendar export
* Monthly Overview export
* Country rollout logic
* Dynamic Programming/Mediaplanning allocation
* Additional daily slot capacity

Pending improvements:

* Monthly Overview formatting refinements
* Additional QA coverage
* UX improvements

---

# Main Workflow

## Data Source

Primary source:

* Airtable

Optional:

* Local JSON export

Main fields include:

* Live Date
* Deadline Delivery
* Property
* Content Type
* YouTube Title
* Description
* Region
* Owner Hub
* Strategy
* TX ID
* TC IN / OUT

---

# Project Structure

## Entry Point

populate_calendar.py

Responsible for:

* Loading Airtable data
* Filtering by month and brand
* Routing export logic
* Calling appropriate exporter

---

## Services

services/planner_service.py

Responsibilities:

* Data loading
* Filtering
* Asset grouping
* Priority sorting
* Shared export helpers

Important functions:

* load_and_filter_records()
* aggregate_by_week_day_and_role()
* assemble_text_for_single_asset()

---

## Exporters

### Disney Channel

exporters/full_calendar.py

Exports:

* FULL CALENDAR YT DC EMEA

### Disney Jr.

exporters/full_calendar_djr.py

Exports:

* FULL CALENDAR YT DJR EMEA

### Monthly Overviews

exporters/monthly_overview.py

exporters/monthly_overview_djr.py

Exports:

* MONTHLY OVERVIEW sheets

---

## Shared

exporters/shared/

Contains:

* Validation
* QA helpers
* Shared export utilities

---

## UI

ui/streamlit_app.py

Provides:

* Brand selector
* Month selector
* QA preview
* Export generation
* Excel download

---

## Country Rebalance Tool

tools/rebalance_country_calendar.py

Purpose:

Apply publishing schedule changes automatically to country calendars.

Current use case:

* Move assets between publishing dates
* Preserve slot structure
* Compact empty slots
* Remove deprecated assets

Supported countries:

* Belgique
* Bulgaria
* Czech Republic
* Denmark
* France
* Greece
* Hungary
* Italy
* MENA
* Netherlands
* Norway
* Poland
* Portugal
* Romania
* Sweden

Israel is currently handled manually due to different requirements.

---

# Important Constraints

## Excel Templates

Do NOT casually modify:

* Merged cells
* Existing formulas
* Sheet structure
* Row positioning

Previous attempts caused significant corruption.

Prefer:

* Additional reports
* Additional sheets
* Validation warnings

instead of dynamically changing template layouts.

---

# QA Philosophy

## Export Blocking Errors

Examples:

* Missing Live Date
* Template not found
* Required sheet missing
* Too many assets for available slots

## Warnings

Examples:

* Missing Region
* Unknown Content Type
* Missing Programming title
* Asset not exported to Monthly Overview

Warnings should not stop export.

---

# Future Roadmap

Priority order:

## 1. Stability

* Prevent regressions
* Improve test coverage
* Reduce duplicated logic

## 2. QA

* Stronger validation
* Better warnings
* Export summaries

## 3. UX

* Better Streamlit experience
* Preview before export
* QA dashboard

## 4. Configuration

Move into config files:

* Content type mappings
* Overview bucket mappings
* Naming rules
* Abbreviations

## 5. Team Usage

Potential future deployment:

* Shared folder
* SharePoint integration
* Multi-user workflow

---

# Development Rules

1. Do not break Disney Channel exports.
2. Do not break Disney Jr. exports.
3. Preserve backwards compatibility whenever possible.
4. Avoid large refactors.
5. Implement changes incrementally.
6. Prefer configuration over hardcoded rules when safe.
7. Explain all architectural changes before implementing them.

---

# Known Future Improvements

* Unplaced Assets report
* Improved Monthly Overview mappings
* Enhanced QA dashboard
* Export history/versioning
* SharePoint integration
* Team-friendly workflows
