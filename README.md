# RosterLine — Shift Allocation App

A small Flask app for managing shift workers and auto-allocating them across
three shifts: **AM** (6am–2pm), **Swing** (2pm–10pm), and **PM** (10pm–6am).

## Features

- **OpenRouter AI Chatbot**: Interactive shift scheduling assistant that allows natural language roster generation, shift swapping, and conversational adjustments.
- **Strict Operational Rules**:
  - **Post-OFF Transition**: Staff coming from an OFF day only enter PM shifts when working.
  - **Baby Care Constraint**: Jacky and Eunice only enter AM and SWING (SW) shifts.
  - **Balanced Workload**: Balanced PM and AM shifts across active staff.
  - **Full Coverage**: Every staff member is allocated a shift or non-working status every day (Shift, OFF, LEAVE, SL, or PH).
- **Public Holiday (PH) & Sick Leave (SL)** options with custom styling and export.
- Add workers one at a time via a form
- Bulk-import workers from an `.xlsx` file (a `Name` column, or just names in
  the first column) — a blank template is available for download in-app
- One-click **Auto-allocate** button that adheres to all operational rules
- Dashboard with a 24-hour timeline strip and a live "now" marker, plus
  color-coded shift columns and live editable weekly grid
- Export the current schedule as a styled `.xlsx` file, PNG image, or PDF
- Remove workers, or clear all shift assignments and re-allocate

## Setup

```bash
python3 -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Then open **http://127.0.0.1:5000** in your browser.

A SQLite database (`instance/shifts.db`) is created automatically on first
run — no separate database setup required.

## Using it

1. Go to **Workers** and add names one by one, or upload an Excel file
   (download the template first if you want the exact expected format).
2. Go back to **Dashboard** and click **Auto-allocate shifts**.
3. Click **Export to Excel** to download the current schedule as a
   spreadsheet you can print or share.

## Project structure

```
app.py                  Flask app: routes, models, allocation & Excel logic
templates/
  base.html              Shared layout, nav, flash messages
  index.html             Dashboard (timeline + shift columns)
  workers.html            Add / import / manage workers
static/
  css/style.css           Styling
  js/main.js              Live "now" marker on the timeline
requirements.txt
```

## Notes on the allocation logic

Auto-allocate shuffles all workers randomly, then assigns them round-robin
across AM → Swing → PM, so shift sizes never differ by more than one person.
Running it again reshuffles from scratch (it doesn't try to keep anyone on
their previous shift).
