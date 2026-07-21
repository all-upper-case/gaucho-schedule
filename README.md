# Gaucho Schedule

A local-first web app for building, checking, preserving, printing, and exporting Gaucho Urbano's weekly employee schedule.

## Features

- Weekly editor grouped by department/role, with autosave and keyboard navigation
- Multiple roles for one employee
- Multiple report times in one role/day, such as `9 / 4`
- Scheduled headcount by department and day
- Blank-cell-to-OFF bulk action
- Publish/lock workflow with downloadable version snapshots
- Automatic preservation of the current week when creating the next week
- Complete SQLite database backups
- Print-friendly schedule, Excel export, and CSV export
- Reviewed POS time-entry CSV import for employee, role, and report-time history
- Windows portable build that does not require Python on the office computer

## Recommended: portable Windows edition

The GitHub Actions workflow named **Build portable Windows app** creates `GauchoSchedule-Windows.zip`.

1. In GitHub, open **Actions** → **Build portable Windows app** → **Run workflow**.
2. When it finishes, download the `GauchoSchedule-Windows` artifact.
3. Extract the ZIP to a normal local folder on the office computer.
4. Double-click `GauchoSchedule.exe`. The schedule opens automatically in the default browser.
5. Leave the small server window open while using the app. Close that window when finished.

Python does not need to be installed on the office computer. The portable app stores its live database in:

```text
%LOCALAPPDATA%\GauchoSchedule\gaucho_schedule.sqlite3
```

The app's **Backup** link downloads a safe copy of the complete database.

## Developer setup

```bat
py -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
py app.py
```

Then open `http://127.0.0.1:5000`.

## Shift entry

- One report time: `3` → `3:00 PM`
- Two report times: `9 / 4` → `9:00 AM / 4:00 PM`
- Not scheduled: `OFF`

End times are intentionally not stored or printed. If an old start-end range is pasted, the app keeps only its start time.

## Importing POS time entries

Open **Import Data**, upload a POS time-entry CSV, and review the proposed employee and job mappings before confirming. The app remembers confirmed mappings, adds roles without removing an employee's existing roles, and learns report-time suggestions by employee, role, and weekday.

Clock-out values are used only during import to distinguish likely meal re-clocks, rapid job transfers, and genuine split shifts. They are never saved as schedule end times or printed. Exact duplicate files are rejected, and every confirmed import is recorded in the import history.

## Schedule preservation

Each week is stored separately. **Create Next Week** publishes, snapshots, and locks the current week before copying it. Published schedules cannot be edited accidentally. Reopening a published week creates another snapshot first. Existing target weeks are never silently overwritten.

## Important security note

This version is intended for local use. It has no login system and must not be exposed directly to the public internet.
