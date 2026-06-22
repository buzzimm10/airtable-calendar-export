# Project Changelog

## 2026-06 — Stable Internal Tooling Version

### Added
- Added Disney Channel Excel export workflow.
- Added Disney Jr. Excel export workflow.
- Added Streamlit local UI for calendar generation.
- Added brand routing for Disney Channel and Disney Jr.
- Added QA mode for basic export validation.
- Added country calendar rebalance automation script.
- Added support for local JSON or Airtable source data.
- Added GitHub repository backup and stable commits.

### Improved
- Refactored project into config, core, exporters, services, tools, and UI modules.
- Improved handling of merged cells in Excel templates.
- Improved owner hub allocation logic.
- Improved Disney Jr. country tier logic.
- Improved content movement logic for country-level calendars.
- Improved slot compaction after moves.

### Notes
- Disney Channel export is considered stable.
- Disney Jr. export is functional but still needs Monthly Overview refinements.
- Country rebalance automation was tested on June 2026 country calendars.
- Israel country calendar is handled manually for now because it has extra TX / TC rows.

### Pending
- Improve QA with blockers, warnings, and info levels.
- Add unknown content type warnings.
- Map Sneak Peek and Videoclip correctly in Monthly Overview.
- Add Unplaced Assets report.
- Improve Streamlit preview and QA cards.
- Add output versioning.
- Move content type rules into config.