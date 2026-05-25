from __future__ import annotations

import csv
import io
import os
import re
import sqlite3
from contextlib import closing
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path

from flask import Flask, Response, flash, redirect, render_template, request, send_file, url_for
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)
DB_PATH = Path(os.environ.get("GAUCHO_SCHEDULE_DB", DATA_DIR / "gaucho_schedule.sqlite3"))

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-local-only-change-before-hosting")

DAYS = ["MON", "TUES", "WEDS", "THURS", "FRI", "SAT", "SUN"]
SHIFT_OPTIONS = [
    "OFF",
    "9:00 AM",
    "10:00 AM",
    "10:30 AM",
    "11:00 AM",
    "12:00 PM",
    "1:00 PM",
    "2:00 PM",
    "2:30 PM",
    "3:00 PM",
    "4:00 PM",
    "5:00 PM",
    "CLOSE",
]

DEFAULT_ROLES = [
    (1, "SERVERS", "(ALCOHOL)"),
    (2, "SERVERS", "(FOOD ONLY)"),
    (3, "GAUCHOS", "(MEAT)"),
    (4, "SERVER", "ASSISTANT"),
    (5, "HOST", ""),
    (6, "KITCHEN", ""),
    (7, "DISHWASHER", ""),
]

DEFAULT_EMPLOYEES = [
    ("EDWARD", 1), ("BRITTANY", 1), ("HUGO", 1), ("JARECK", 1), ("ERICK", 1),
    ("ROMER", 1), ("DANIEL P.", 1), ("JORDAN", 1), ("SAM", 1),
    ("MIGUEL", 2), ("DANIEL", 2), ("DIANA", 2), ("MICHAEL", 2),
    ("REYNALDO", 3), ("DAVID M.", 3), ("JUAN", 3), ("ENRIQUE", 3), ("NICOLAS", 3),
    ("WILLIAM", 3), ("KEVIN", 3), ("JAIDER", 3), ("LUCAS", 3), ("ISABEL", 3),
    ("YERSON", 3), ("JOEL", 3), ("JULIO", 3),
    ("MARTIN", 4), ("KISHANA", 4), ("LEXI", 4),
    ("EMILIO", 5), ("NICOLE", 5), ("ERIKA", 5), ("JEAN", 5),
    ("ROBERTO", 6), ("SVETLANA", 6), ("JARECK", 6), ("RAFAEL", 6), ("JESUS", 6),
    ("MICHAEL", 6), ("CINTIA", 6), ("ANATOLE", 6), ("DANIEL", 6), ("JESSICA", 6), ("ALEXIS", 6),
    ("ROBIN", 7), ("JONH", 7), ("FRANKLIN", 7), ("JOSETH", 7),
]


@dataclass
class Role:
    id: int
    title: str
    subtitle: str
    display_order: int


@dataclass
class Employee:
    id: int
    name: str
    role_id: int
    active: int
    display_order: int


def get_db() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db() -> None:
    with closing(get_db()) as db:
        db.executescript(
            """
            CREATE TABLE IF NOT EXISTS roles (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT NOT NULL,
                subtitle TEXT NOT NULL DEFAULT '',
                display_order INTEGER NOT NULL DEFAULT 0
            );

            CREATE TABLE IF NOT EXISTS employees (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                role_id INTEGER NOT NULL REFERENCES roles(id),
                active INTEGER NOT NULL DEFAULT 1,
                display_order INTEGER NOT NULL DEFAULT 0
            );

            CREATE TABLE IF NOT EXISTS weeks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                start_date TEXT NOT NULL UNIQUE,
                notes TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS shifts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                week_id INTEGER NOT NULL REFERENCES weeks(id) ON DELETE CASCADE,
                employee_id INTEGER NOT NULL REFERENCES employees(id) ON DELETE CASCADE,
                day_index INTEGER NOT NULL CHECK(day_index BETWEEN 0 AND 6),
                label TEXT NOT NULL DEFAULT 'OFF',
                UNIQUE(week_id, employee_id, day_index)
            );
            """
        )

        role_count = db.execute("SELECT COUNT(*) FROM roles").fetchone()[0]
        if role_count == 0:
            db.executemany(
                "INSERT INTO roles (display_order, title, subtitle) VALUES (?, ?, ?)",
                [(order, title, subtitle) for order, title, subtitle in DEFAULT_ROLES],
            )

        employee_count = db.execute("SELECT COUNT(*) FROM employees").fetchone()[0]
        if employee_count == 0:
            counters: dict[int, int] = {}
            rows = []
            for name, role_id in DEFAULT_EMPLOYEES:
                counters[role_id] = counters.get(role_id, 0) + 1
                rows.append((name, role_id, 1, counters[role_id]))
            db.executemany(
                "INSERT INTO employees (name, role_id, active, display_order) VALUES (?, ?, ?, ?)",
                rows,
            )
        db.commit()


def monday_for(value: date | None = None) -> date:
    value = value or date.today()
    return value - timedelta(days=value.weekday())


def parse_week_start(raw: str | None) -> date:
    if not raw:
        return monday_for()
    parsed = datetime.strptime(raw, "%Y-%m-%d").date()
    return monday_for(parsed)


def get_or_create_week(db: sqlite3.Connection, week_start: date) -> int:
    start = week_start.isoformat()
    row = db.execute("SELECT id FROM weeks WHERE start_date = ?", (start,)).fetchone()
    if row:
        return int(row["id"])
    cur = db.execute("INSERT INTO weeks (start_date) VALUES (?)", (start,))
    db.commit()
    return int(cur.lastrowid)


def roles(db: sqlite3.Connection) -> list[Role]:
    return [Role(**dict(r)) for r in db.execute("SELECT * FROM roles ORDER BY display_order, id")]


def employees(db: sqlite3.Connection, include_inactive: bool = False) -> list[Employee]:
    where = "" if include_inactive else "WHERE active = 1"
    return [
        Employee(**dict(r))
        for r in db.execute(f"SELECT * FROM employees {where} ORDER BY role_id, display_order, name, id")
    ]


def shift_map(db: sqlite3.Connection, week_id: int) -> dict[tuple[int, int], str]:
    rows = db.execute("SELECT employee_id, day_index, label FROM shifts WHERE week_id = ?", (week_id,)).fetchall()
    return {(int(r["employee_id"]), int(r["day_index"])): r["label"] for r in rows}


def grouped_schedule(db: sqlite3.Connection, week_start: date):
    week_id = get_or_create_week(db, week_start)
    all_roles = roles(db)
    all_employees = employees(db)
    shifts = shift_map(db, week_id)
    by_role: dict[int, list[Employee]] = {role.id: [] for role in all_roles}
    for emp in all_employees:
        by_role.setdefault(emp.role_id, []).append(emp)
    return week_id, [(role, by_role.get(role.id, [])) for role in all_roles], shifts


def validate_shift(label: str) -> list[str]:
    value = label.strip().upper()
    warnings: list[str] = []
    if not value:
        return warnings
    if value in {"OFF", "CLOSE"}:
        return warnings
    if ";" in value:
        warnings.append("uses a semicolon; consider changing it to a clear start/end time like 10:00 AM-4:00 PM")
    if re.fullmatch(r"\d{3,4}", value):
        warnings.append("looks like 24-hour time; consider writing it as 5:00 PM, 1:00 PM, etc.")
    if value == "12:00 AM":
        warnings.append("is midnight; confirm this is intentional and not noon/closing")
    if re.fullmatch(r"\d{1,2}:\d{2}\s*(AM|PM)", value):
        minute = int(value.split(":", 1)[1][:2])
        if minute not in {0, 30}:
            warnings.append("has unusual minutes; confirm this is not a typo")
    return warnings


def collect_warnings(db: sqlite3.Connection, week_id: int) -> list[str]:
    rows = db.execute(
        """
        SELECT employees.name, shifts.day_index, shifts.label
        FROM shifts
        JOIN employees ON employees.id = shifts.employee_id
        WHERE shifts.week_id = ? AND TRIM(shifts.label) != '' AND UPPER(TRIM(shifts.label)) != 'OFF'
        ORDER BY employees.name, shifts.day_index
        """,
        (week_id,),
    ).fetchall()
    warnings: list[str] = []
    for row in rows:
        for warning in validate_shift(row["label"]):
            warnings.append(f"{row['name']} on {DAYS[int(row['day_index'])]}: {row['label']} {warning}.")
    return warnings


def week_dates(week_start: date) -> list[date]:
    return [week_start + timedelta(days=i) for i in range(7)]


@app.before_request
def ensure_db() -> None:
    init_db()


@app.route("/")
def index():
    with closing(get_db()) as db:
        week_rows = db.execute("SELECT * FROM weeks ORDER BY start_date DESC LIMIT 12").fetchall()
    return render_template("index.html", weeks=week_rows, current_week=monday_for())


@app.route("/week", methods=["POST"])
def go_to_week():
    week_start = parse_week_start(request.form.get("week_start"))
    return redirect(url_for("edit_week", week_start=week_start.isoformat()))


@app.route("/week/<week_start>")
def edit_week(week_start: str):
    start = parse_week_start(week_start)
    with closing(get_db()) as db:
        week_id, grouped, shifts = grouped_schedule(db, start)
        warnings = collect_warnings(db, week_id)
    return render_template(
        "schedule.html",
        week_start=start,
        week_end=start + timedelta(days=6),
        dates=week_dates(start),
        days=DAYS,
        grouped=grouped,
        shifts=shifts,
        shift_options=SHIFT_OPTIONS,
        warnings=warnings,
    )


@app.route("/week/<week_start>/save", methods=["POST"])
def save_week(week_start: str):
    start = parse_week_start(week_start)
    with closing(get_db()) as db:
        week_id = get_or_create_week(db, start)
        active_employees = employees(db)
        for emp in active_employees:
            for day_index in range(7):
                key = f"shift_{emp.id}_{day_index}"
                label = request.form.get(key, "OFF").strip() or "OFF"
                db.execute(
                    """
                    INSERT INTO shifts (week_id, employee_id, day_index, label)
                    VALUES (?, ?, ?, ?)
                    ON CONFLICT(week_id, employee_id, day_index)
                    DO UPDATE SET label = excluded.label
                    """,
                    (week_id, emp.id, day_index, label),
                )
        db.commit()
        warning_count = len(collect_warnings(db, week_id))
    flash(f"Schedule saved. {warning_count} warning(s) found." if warning_count else "Schedule saved.")
    return redirect(url_for("edit_week", week_start=start.isoformat()))


@app.route("/week/<week_start>/copy-next", methods=["POST"])
def copy_next_week(week_start: str):
    start = parse_week_start(week_start)
    next_start = start + timedelta(days=7)
    with closing(get_db()) as db:
        source_id = get_or_create_week(db, start)
        target_id = get_or_create_week(db, next_start)
        db.execute("DELETE FROM shifts WHERE week_id = ?", (target_id,))
        db.execute(
            """
            INSERT INTO shifts (week_id, employee_id, day_index, label)
            SELECT ?, employee_id, day_index, label FROM shifts WHERE week_id = ?
            """,
            (target_id, source_id),
        )
        db.commit()
    flash("Next week created by copying the current week.")
    return redirect(url_for("edit_week", week_start=next_start.isoformat()))


@app.route("/print/<week_start>")
def print_week(week_start: str):
    start = parse_week_start(week_start)
    with closing(get_db()) as db:
        _, grouped, shifts = grouped_schedule(db, start)
    return render_template(
        "print.html",
        week_start=start,
        week_end=start + timedelta(days=6),
        dates=week_dates(start),
        days=DAYS,
        grouped=grouped,
        shifts=shifts,
    )


@app.route("/employees", methods=["GET", "POST"])
def manage_employees():
    with closing(get_db()) as db:
        if request.method == "POST":
            action = request.form.get("action")
            if action == "add":
                name = request.form.get("name", "").strip().upper()
                role_id = int(request.form.get("role_id", "1"))
                max_order = db.execute("SELECT COALESCE(MAX(display_order), 0) FROM employees WHERE role_id = ?", (role_id,)).fetchone()[0]
                if name:
                    db.execute(
                        "INSERT INTO employees (name, role_id, active, display_order) VALUES (?, ?, 1, ?)",
                        (name, role_id, int(max_order) + 1),
                    )
                    db.commit()
                    flash(f"Added {name}.")
            elif action == "update":
                emp_id = int(request.form["employee_id"])
                name = request.form.get("name", "").strip().upper()
                role_id = int(request.form.get("role_id", "1"))
                active = 1 if request.form.get("active") == "on" else 0
                display_order = int(request.form.get("display_order", "0") or 0)
                db.execute(
                    "UPDATE employees SET name = ?, role_id = ?, active = ?, display_order = ? WHERE id = ?",
                    (name, role_id, active, display_order, emp_id),
                )
                db.commit()
                flash("Employee updated.")
            return redirect(url_for("manage_employees"))
        all_roles = roles(db)
        all_employees = employees(db, include_inactive=True)
    return render_template("employees.html", roles=all_roles, employees=all_employees)


@app.route("/export/<week_start>.csv")
def export_csv(week_start: str):
    start = parse_week_start(week_start)
    with closing(get_db()) as db:
        _, grouped, shifts = grouped_schedule(db, start)
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["GAUCHO URBANO EMPLOYEE SCHEDULE", start.isoformat(), "through", (start + timedelta(days=6)).isoformat()])
    for role, emps in grouped:
        writer.writerow([])
        writer.writerow([role.title, role.subtitle, *DAYS])
        writer.writerow(["", "", *[d.day for d in week_dates(start)]])
        for emp in emps:
            writer.writerow([emp.name, "", *[shifts.get((emp.id, i), "OFF") for i in range(7)]])
    return Response(
        output.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": f"attachment; filename=gaucho-schedule-{start.isoformat()}.csv"},
    )


@app.route("/export/<week_start>.xlsx")
def export_xlsx(week_start: str):
    start = parse_week_start(week_start)
    with closing(get_db()) as db:
        _, grouped, shifts = grouped_schedule(db, start)
    wb = Workbook()
    ws = wb.active
    ws.title = start.strftime("%b %d")
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    row = 1
    ws.cell(row=row, column=1, value="GAUCHO URBANO EMPLOYEE SCHEDULE").font = Font(bold=True)
    ws.cell(row=row, column=5, value=start.strftime("%-m/%-d/%Y") if os.name != "nt" else start.strftime("%#m/%#d/%Y"))
    ws.cell(row=row, column=6, value="through")
    ws.cell(row=row, column=7, value=(start + timedelta(days=6)).strftime("%-m/%-d/%Y") if os.name != "nt" else (start + timedelta(days=6)).strftime("%#m/%#d/%Y"))
    row += 1
    for role, emps in grouped:
        row += 1
        ws.cell(row=row, column=1, value=role.title).font = Font(bold=True)
        if role.subtitle:
            ws.cell(row=row + 1, column=1, value=role.subtitle).font = Font(bold=True)
        for idx, day in enumerate(DAYS, start=3):
            ws.cell(row=row, column=idx, value=day).font = Font(bold=True, italic=True)
            ws.cell(row=row + 1, column=idx, value=week_dates(start)[idx - 3].day).font = Font(bold=True)
        row += 2
        for emp in emps:
            ws.cell(row=row, column=1, value=emp.name)
            for day_index in range(7):
                ws.cell(row=row, column=day_index + 3, value=shifts.get((emp.id, day_index), "OFF"))
            for col in range(1, 10):
                ws.cell(row=row, column=col).border = border
                ws.cell(row=row, column=col).alignment = Alignment(horizontal="center")
            ws.cell(row=row, column=1).alignment = Alignment(horizontal="left")
            row += 1
        row += 1
    for col in range(1, 10):
        ws.column_dimensions[get_column_letter(col)].width = 14 if col > 1 else 18
    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)
    return send_file(
        file_stream,
        as_attachment=True,
        download_name=f"gaucho-schedule-{start.isoformat()}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    init_db()
    default_host = "0.0.0.0" if os.environ.get("REPL_ID") else "127.0.0.1"
    host = os.environ.get("GAUCHO_SCHEDULE_HOST", default_host)
    port = int(os.environ.get("PORT", "5000"))
    app.run(host=host, port=port, debug=True)
