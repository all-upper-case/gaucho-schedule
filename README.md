# Gaucho Schedule

A small local-first web app for editing and printing Gaucho Urbano's weekly employee schedule.

The app is designed to run on the office computer at `http://127.0.0.1:5000` by default. It can also be tested in Replit and later deployed to a small web host if phone/multi-manager access becomes useful.

## Current MVP features

- Weekly schedule editor grouped by department/role
- Seed employee list based on the current Excel/PDF schedule layout
- Copy current week into the next week
- Employee add/edit/deactivate screen
- Print-friendly schedule view
- Excel (`.xlsx`) and CSV export
- Basic warnings for suspicious shift entries, such as `10;1600`, `1700`, odd minutes, or `12:00 AM`

## Run locally on Windows

Open **Command Prompt** in the project folder and run:

```bat
py -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
py app.py
```

Then open this address in the browser:

```text
http://127.0.0.1:5000
```

To stop the app, click the Command Prompt window and press `Ctrl+C`.

## Run in Replit

Create/import a Replit project from this GitHub repository. For a basic Python Repl, set the run command to:

```bash
python app.py
```

Replit may expose the app through its own web preview instead of `127.0.0.1:5000`.

## Data storage

By default, the SQLite database is created here:

```text
data/gaucho_schedule.sqlite3
```

That database file is intentionally ignored by Git so employee/schedule data does not get pushed to the public repository.

To use a different database path, set the `GAUCHO_SCHEDULE_DB` environment variable.

## Important security note

This first version is intended for local use or private testing. It does not include login/password protection yet. Do not deploy it publicly for real manager use until authentication is added.
