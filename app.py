from __future__ import annotations

import csv
import io
import json
import os
import re
import sqlite3
import sys
import threading
import webbrowser
from contextlib import closing
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from pathlib import Path
from tempfile import NamedTemporaryFile

from flask import Flask, Response, flash, jsonify, redirect, render_template, request, send_file, url_for
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


SOURCE_DIR = Path(__file__).resolve().parent
RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", SOURCE_DIR))
IS_PACKAGED = bool(getattr(sys, "frozen", False))

if os.environ.get("GAUCHO_SCHEDULE_DATA_DIR"):
    DATA_DIR = Path(os.environ["GAUCHO_SCHEDULE_DATA_DIR"])
elif IS_PACKAGED:
    local_app_data = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
    DATA_DIR = local_app_data / "GauchoSchedule"
else:
    DATA_DIR = SOURCE_DIR / "data"
DATA_DIR.mkdir(parents=True, exist_ok=True)
DB_PATH = Path(os.environ.get("GAUCHO_SCHEDULE_DB", DATA_DIR / "gaucho_schedule.sqlite3"))

app = Flask(
    __name__,
    template_folder=str(RESOURCE_DIR / "templates"),
    static_folder=str(RESOURCE_DIR / "static"),
)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-local-only-change-before-hosting")

DAYS = ["MON", "TUES", "WEDS", "THURS", "FRI", "SAT", "SUN"]
DAY_WORDS = set(DAYS + ["THU"])
GLOBAL_SHIFT_OPTIONS = [
    "9:00 AM-12:00 PM / 4:00 PM-9:00 PM",
    "10:00 AM-4:00 PM",
    "3:00 PM-9:00 PM",
    "4:00 PM-10:00 PM",
    "3:00 PM",
    "4:00 PM",
    "5:00 PM",
    "CLOSE",
    "OFF",
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
    ("EDWARD", 1), ("BRITTANY", 1), ("HUGO", 1), ("JARECK", 1), ("ERICK", 1), ("ROMER", 1), ("DANIEL P.", 1), ("JORDAN", 1), ("SAM", 1),
    ("MIGUEL", 2), ("DANIEL", 2), ("DIANA", 2), ("MICHAEL", 2),
    ("REYNALDO", 3), ("DAVID M.", 3), ("JUAN", 3), ("ENRIQUE", 3), ("NICOLAS", 3), ("WILLIAM", 3), ("KEVIN", 3), ("JAIDER", 3), ("LUCAS", 3), ("ISABEL", 3), ("YERSON", 3), ("JOEL", 3), ("JULIO", 3),
    ("MARTIN", 4), ("KISHANA", 4), ("LEXI", 4),
    ("EMILIO", 5), ("NICOLE", 5), ("ERIKA", 5), ("JEAN", 5),
    ("ROBERTO", 6), ("SVETLANA", 6), ("JARECK", 6), ("RAFAEL", 6), ("JESUS", 6), ("MICHAEL", 6), ("CINTIA", 6), ("ANATOLE", 6), ("DANIEL", 6), ("JESSICA", 6), ("ALEXIS", 6),
    ("ROBIN", 7), ("JONH", 7), ("FRANKLIN", 7), ("JOSETH", 7),
]


@dataclass
class Role:
    id: int
    title: str
    subtitle: str
    display_order: int


@dataclass
class ScheduleEmployee:
    assignment_id: int
    employee_id: int
    name: str
    role_id: int
    date_of_birth: str | None
    parental_late_consent: int
    school_start_time: str
    school_end_time: str
    display_order: int
    archived_at: str | None = None


def get_db() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA journal_mode = WAL")
    return conn


def table_exists(db: sqlite3.Connection, table: str) -> bool:
    return db.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (table,)).fetchone() is not None


def ensure_column(db: sqlite3.Connection, table: str, column: str, ddl: str) -> None:
    cols = {row["name"] for row in db.execute(f"PRAGMA table_info({table})")}
    if column not in cols:
        db.execute(f"ALTER TABLE {table} ADD COLUMN {column} {ddl}")


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
                display_order INTEGER NOT NULL DEFAULT 0,
                archived_at TEXT
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
                label TEXT NOT NULL DEFAULT '',
                UNIQUE(week_id, employee_id, day_index)
            );
            CREATE TABLE IF NOT EXISTS shift_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                role_title TEXT NOT NULL,
                role_subtitle TEXT NOT NULL DEFAULT '',
                employee_name TEXT NOT NULL,
                day_index INTEGER NOT NULL CHECK(day_index BETWEEN 0 AND 6),
                label TEXT NOT NULL,
                source_sheet TEXT NOT NULL DEFAULT '',
                count INTEGER NOT NULL DEFAULT 1,
                UNIQUE(role_title, role_subtitle, employee_name, day_index, label, source_sheet)
            );
            CREATE TABLE IF NOT EXISTS employee_assignments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id INTEGER NOT NULL REFERENCES employees(id) ON DELETE CASCADE,
                role_id INTEGER NOT NULL REFERENCES roles(id),
                active INTEGER NOT NULL DEFAULT 1,
                display_order INTEGER NOT NULL DEFAULT 0,
                archived_at TEXT,
                UNIQUE(employee_id, role_id)
            );
            CREATE TABLE IF NOT EXISTS schedule_entries (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                week_id INTEGER NOT NULL REFERENCES weeks(id) ON DELETE CASCADE,
                assignment_id INTEGER NOT NULL REFERENCES employee_assignments(id) ON DELETE CASCADE,
                day_index INTEGER NOT NULL CHECK(day_index BETWEEN 0 AND 6),
                label TEXT NOT NULL DEFAULT '',
                UNIQUE(week_id, assignment_id, day_index)
            );
            CREATE TABLE IF NOT EXISTS week_versions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                week_id INTEGER NOT NULL REFERENCES weeks(id) ON DELETE CASCADE,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                reason TEXT NOT NULL,
                snapshot_json TEXT NOT NULL
            );
            """
        )
        ensure_column(db, "employees", "archived_at", "TEXT")
        ensure_column(db, "employees", "date_of_birth", "TEXT")
        ensure_column(db, "employees", "parental_late_consent", "INTEGER NOT NULL DEFAULT 0")
        ensure_column(db, "employees", "school_start_time", "TEXT NOT NULL DEFAULT '08:00'")
        ensure_column(db, "employees", "school_end_time", "TEXT NOT NULL DEFAULT '15:00'")
        ensure_column(db, "weeks", "school_in_session", "INTEGER NOT NULL DEFAULT 1")
        ensure_column(db, "weeks", "school_day_mask", "TEXT NOT NULL DEFAULT '1111100'")
        ensure_column(db, "weeks", "published_at", "TEXT")

        if db.execute("SELECT COUNT(*) FROM roles").fetchone()[0] == 0:
            db.executemany(
                "INSERT INTO roles (display_order, title, subtitle) VALUES (?, ?, ?)",
                [(order, title, subtitle) for order, title, subtitle in DEFAULT_ROLES],
            )
        if db.execute("SELECT COUNT(*) FROM employees").fetchone()[0] == 0:
            counters: dict[int, int] = {}
            rows = []
            for name, role_id in DEFAULT_EMPLOYEES:
                counters[role_id] = counters.get(role_id, 0) + 1
                rows.append((name, role_id, 1, counters[role_id]))
            db.executemany(
                "INSERT INTO employees (name, role_id, active, display_order) VALUES (?, ?, ?, ?)", rows
            )

        if db.execute("SELECT COUNT(*) FROM employee_assignments").fetchone()[0] == 0:
            db.execute(
                """INSERT OR IGNORE INTO employee_assignments
                   (employee_id, role_id, active, display_order, archived_at)
                   SELECT id, role_id, active, display_order, archived_at FROM employees"""
            )
        if db.execute("SELECT COUNT(*) FROM schedule_entries").fetchone()[0] == 0 and table_exists(db, "shifts"):
            db.execute(
                """INSERT OR IGNORE INTO schedule_entries (week_id, assignment_id, day_index, label)
                   SELECT s.week_id, a.id, s.day_index, s.label
                   FROM shifts s
                   JOIN employees e ON e.id = s.employee_id
                   JOIN employee_assignments a ON a.employee_id = e.id AND a.role_id = e.role_id"""
            )
        db.commit()


def monday_for(value: date | None = None) -> date:
    value = value or date.today()
    return value - timedelta(days=value.weekday())


def parse_week_start(raw: str | None) -> date:
    if not raw:
        return monday_for()
    return monday_for(datetime.strptime(raw, "%Y-%m-%d").date())


def get_or_create_week(db: sqlite3.Connection, week_start: date) -> int:
    row = db.execute("SELECT id FROM weeks WHERE start_date = ?", (week_start.isoformat(),)).fetchone()
    if row:
        return int(row["id"])
    cur = db.execute("INSERT INTO weeks (start_date) VALUES (?)", (week_start.isoformat(),))
    db.commit()
    return int(cur.lastrowid)


def week_row(db: sqlite3.Connection, week_start: date) -> sqlite3.Row:
    week_id = get_or_create_week(db, week_start)
    return db.execute("SELECT * FROM weeks WHERE id = ?", (week_id,)).fetchone()


def is_week_locked(db: sqlite3.Connection, week_id: int) -> bool:
    return db.execute("SELECT published_at FROM weeks WHERE id = ?", (week_id,)).fetchone()[0] is not None


def roles(db: sqlite3.Connection) -> list[Role]:
    return [Role(**dict(row)) for row in db.execute("SELECT * FROM roles ORDER BY display_order, id")]


def schedule_employees(db: sqlite3.Connection, archived: bool = False) -> list[ScheduleEmployee]:
    where = "a.archived_at IS NOT NULL" if archived else "a.active = 1 AND a.archived_at IS NULL AND e.active = 1"
    rows = db.execute(
        f"""SELECT a.id AS assignment_id, e.id AS employee_id, e.name, a.role_id,
                   e.date_of_birth, e.parental_late_consent, e.school_start_time,
                   e.school_end_time, a.display_order, a.archived_at
            FROM employee_assignments a
            JOIN employees e ON e.id = a.employee_id
            WHERE {where}
            ORDER BY a.role_id, a.display_order, e.name, a.id"""
    )
    return [ScheduleEmployee(**dict(row)) for row in rows]


def grouped_employees(db: sqlite3.Connection, archived: bool = False):
    all_roles = roles(db)
    by_role: dict[int, list[ScheduleEmployee]] = {role.id: [] for role in all_roles}
    for employee in schedule_employees(db, archived=archived):
        by_role.setdefault(employee.role_id, []).append(employee)
    return [(role, by_role.get(role.id, [])) for role in all_roles]


def assignment_lookup(db: sqlite3.Connection) -> dict[int, ScheduleEmployee]:
    return {employee.assignment_id: employee for employee in schedule_employees(db)}


def shift_map(db: sqlite3.Connection, week_id: int) -> dict[tuple[int, int], str]:
    rows = db.execute(
        "SELECT assignment_id, day_index, label FROM schedule_entries WHERE week_id = ?", (week_id,)
    ).fetchall()
    return {(int(row["assignment_id"]), int(row["day_index"])): row["label"] for row in rows}


def grouped_schedule(db: sqlite3.Connection, week_start: date):
    week = week_row(db, week_start)
    return week, grouped_employees(db), shift_map(db, int(week["id"]))


def time_to_12h(hour: int, minute: int = 0) -> str:
    hour %= 24
    return f"{hour % 12 or 12}:{minute:02d} {'AM' if hour < 12 else 'PM'}"


def format_minutes(total_minutes: int, output_format: str = "12h") -> str:
    total_minutes %= 24 * 60
    hour, minute = divmod(total_minutes, 60)
    return f"{hour:02d}:{minute:02d}" if output_format == "24h" else time_to_12h(hour, minute)


def normalize_one_time(value, output_format: str = "12h") -> str:
    raw = str(value or "").strip()
    if not raw:
        return ""
    upper = raw.upper()
    if upper in {"CL", "CLOSE", "CLOSING"}:
        return "CLOSE"
    compact = re.fullmatch(r"\d{3,4}", raw)
    if compact:
        hour, minute = int(raw[:-2]), int(raw[-2:])
        if hour <= 23 and minute < 60:
            if hour <= 7:
                hour += 12
            return format_minutes(hour * 60 + minute, output_format)
    match = re.fullmatch(r"(\d{1,2})(?::(\d{1,2}))?\s*([AaPp][Mm])?", raw.replace(".", ":"))
    if match:
        hour = int(match.group(1))
        minute = int(match.group(2) or 0)
        suffix = match.group(3).upper() if match.group(3) else None
        if minute >= 60:
            return upper
        if suffix == "PM" and hour < 12:
            hour += 12
        elif suffix == "AM" and hour == 12:
            hour = 0
        elif not suffix and hour <= 7:
            hour += 12
        if hour <= 23:
            return format_minutes(hour * 60 + minute, output_format)
    return upper


def normalize_shift_label(value, output_format: str = "12h") -> str:
    if value is None or isinstance(value, bool):
        return ""
    if isinstance(value, time):
        return format_minutes(value.hour * 60 + value.minute, output_format)
    if isinstance(value, datetime):
        return format_minutes(value.hour * 60 + value.minute, output_format) if value.hour or value.minute else ""
    if isinstance(value, (int, float)):
        if float(value) == 0:
            return "OFF"
        if 0 < float(value) < 1:
            return format_minutes(round(float(value) * 24 * 60), output_format)
    raw = str(value).strip()
    if not raw:
        return ""
    upper = raw.upper()
    if upper in {"0", "OFF", "OOF"}:
        return "OFF"
    if upper in {"CL", "CLOSE", "CLOSING"}:
        return "CLOSE"

    multi_parts = [part.strip() for part in re.split(r"\s*(?:/|,|\s+&\s+)\s*", raw) if part.strip()]
    normalized_parts: list[str] = []
    for part in multi_parts:
        range_parts = [piece.strip() for piece in re.split(r"\s*(?:;|\s[-–—]\s|(?<=\d)[-–—](?=\d))\s*", part) if piece.strip()]
        if len(range_parts) == 2:
            normalized_start = normalize_one_time(range_parts[0], output_format)
            normalized_end = normalize_one_time(range_parts[1], output_format)
            start_minutes = parse_time_minutes(normalized_start)
            end_minutes = parse_time_minutes(normalized_end)
            end_has_suffix = bool(re.search(r"[AaPp][Mm]", range_parts[1]))
            if not end_has_suffix and start_minutes is not None and end_minutes is not None and end_minutes <= start_minutes:
                normalized_end = format_minutes(end_minutes + 12 * 60, output_format)
            normalized_parts.append(f"{normalized_start}-{normalized_end}")
        else:
            normalized_parts.append(normalize_one_time(part, output_format))
    return " / ".join(normalized_parts)


def parse_time_minutes(raw: str) -> int | None:
    value = raw.strip().upper()
    if value == "CLOSE":
        return None
    match = re.fullmatch(r"(\d{1,2}):(\d{2})\s*([AP]M)?", value)
    if not match:
        return None
    hour, minute = int(match.group(1)), int(match.group(2))
    suffix = match.group(3)
    if suffix == "PM" and hour < 12:
        hour += 12
    if suffix == "AM" and hour == 12:
        hour = 0
    return hour * 60 + minute


def shift_intervals(label: str) -> tuple[list[tuple[int, int]], bool]:
    if not label or label == "OFF":
        return [], True
    intervals: list[tuple[int, int]] = []
    complete = True
    for part in [piece.strip() for piece in label.split("/") if piece.strip()]:
        pieces = [piece.strip() for piece in part.split("-", 1)]
        if len(pieces) != 2:
            complete = False
            continue
        start, end = parse_time_minutes(pieces[0]), parse_time_minutes(pieces[1])
        if start is None or end is None:
            complete = False
            continue
        if end <= start:
            end += 24 * 60
        intervals.append((start, end))
    return intervals, complete


def merge_intervals(intervals: list[tuple[int, int]]) -> list[tuple[int, int]]:
    merged: list[list[int]] = []
    for start, end in sorted(intervals):
        if not merged or start > merged[-1][1]:
            merged.append([start, end])
        else:
            merged[-1][1] = max(merged[-1][1], end)
    return [(start, end) for start, end in merged]


def age_on(dob_raw: str | None, day: date) -> int | None:
    if not dob_raw:
        return None
    try:
        dob = date.fromisoformat(dob_raw)
    except ValueError:
        return None
    return day.year - dob.year - ((day.month, day.day) < (dob.month, dob.day))


def label_is_usable(label: str, include_off: bool = True) -> bool:
    if not label:
        return False
    if label == "OFF":
        return include_off
    if label == "CLOSE":
        return True
    return bool(re.search(r"\d{1,2}:\d{2}", label))


def unique_options(*groups: list[str]) -> list[str]:
    seen: set[str] = set()
    output: list[str] = []
    for group in groups:
        for item in group:
            clean = (item or "").strip()
            if clean and clean not in seen:
                seen.add(clean)
                output.append(clean)
    return output


def previous_week_options(db: sqlite3.Connection, week_start: date) -> dict[tuple[int, int], str]:
    row = db.execute("SELECT id FROM weeks WHERE start_date = ?", ((week_start - timedelta(days=7)).isoformat(),)).fetchone()
    return shift_map(db, int(row["id"])) if row else {}


def history_options_for_cell(db: sqlite3.Connection, role: Role, employee: ScheduleEmployee, day_index: int) -> list[str]:
    rows = db.execute(
        """SELECT label, SUM(count) AS total FROM shift_history
           WHERE role_title = ? AND role_subtitle = ? AND employee_name = ? AND day_index = ? AND label != 'OFF'
           GROUP BY label ORDER BY total DESC, label LIMIT 12""",
        (role.title, role.subtitle, employee.name.upper(), day_index),
    ).fetchall()
    return [row["label"] for row in rows]


def role_options_for_cell(db: sqlite3.Connection, role: Role, day_index: int) -> list[str]:
    rows = db.execute(
        """SELECT label, SUM(count) AS total FROM shift_history
           WHERE role_title = ? AND role_subtitle = ? AND day_index = ? AND label != 'OFF'
           GROUP BY label ORDER BY total DESC, label LIMIT 10""",
        (role.title, role.subtitle, day_index),
    ).fetchall()
    return [row["label"] for row in rows]


def global_options_for_cell(db: sqlite3.Connection, day_index: int) -> list[str]:
    rows = db.execute(
        """SELECT label, SUM(count) AS total FROM shift_history
           WHERE day_index = ? AND label != 'OFF'
           GROUP BY label ORDER BY total DESC, label LIMIT 10""",
        (day_index,),
    ).fetchall()
    return [row["label"] for row in rows]


def best_label_for_cell(db: sqlite3.Connection, role: Role, employee: ScheduleEmployee, day_index: int) -> str:
    sources = [
        history_options_for_cell(db, role, employee, day_index),
        role_options_for_cell(db, role, day_index),
        global_options_for_cell(db, day_index),
        GLOBAL_SHIFT_OPTIONS,
    ]
    for source in sources:
        for label in source:
            if label and label != "OFF":
                return label
    return "OFF"


def build_suggestions(db: sqlite3.Connection, grouped, week_start: date) -> dict[tuple[int, int], list[str]]:
    previous = previous_week_options(db, week_start)
    suggestions: dict[tuple[int, int], list[str]] = {}
    for role, employees_for_role in grouped:
        for employee in employees_for_role:
            for day_index in range(7):
                prev = previous.get((employee.assignment_id, day_index))
                suggestions[(employee.assignment_id, day_index)] = unique_options(
                    [prev] if prev else [],
                    history_options_for_cell(db, role, employee, day_index),
                    role_options_for_cell(db, role, day_index),
                    global_options_for_cell(db, day_index),
                    GLOBAL_SHIFT_OPTIONS,
                )[:20]
    return suggestions


def save_shift_value(db: sqlite3.Connection, week_id: int, employee: ScheduleEmployee, role: Role, day_index: int, raw_label: str) -> str:
    label = normalize_shift_label(raw_label or "", "12h")
    db.execute(
        """INSERT INTO schedule_entries (week_id, assignment_id, day_index, label)
           VALUES (?, ?, ?, ?)
           ON CONFLICT(week_id, assignment_id, day_index) DO UPDATE SET label = excluded.label""",
        (week_id, employee.assignment_id, day_index, label),
    )
    if label_is_usable(label):
        db.execute(
            """INSERT INTO shift_history
               (role_title, role_subtitle, employee_name, day_index, label, source_sheet, count)
               VALUES (?, ?, ?, ?, ?, 'manual-app-entry', 1)
               ON CONFLICT(role_title, role_subtitle, employee_name, day_index, label, source_sheet)
               DO UPDATE SET count = count + 1""",
            (role.title, role.subtitle, employee.name.upper(), day_index, label),
        )
    return label


def school_days_for_week(week: sqlite3.Row) -> set[int]:
    if not int(week["school_in_session"]):
        return set()
    mask = (week["school_day_mask"] or "1111100").ljust(7, "0")
    return {index for index, value in enumerate(mask[:7]) if value == "1"}


def is_summer_evening(day: date) -> bool:
    labor_day = date(day.year, 9, 1)
    while labor_day.weekday() != 0:
        labor_day += timedelta(days=1)
    return date(day.year, 6, 1) <= day <= labor_day


def collect_warnings(db: sqlite3.Connection, week: sqlite3.Row, week_start: date) -> list[dict[str, str]]:
    rows = db.execute(
        """SELECT e.id AS employee_id, e.name, e.date_of_birth, e.parental_late_consent,
                  e.school_start_time, e.school_end_time, a.role_id, r.title AS role_title,
                  se.day_index, se.label
           FROM schedule_entries se
           JOIN employee_assignments a ON a.id = se.assignment_id
           JOIN employees e ON e.id = a.employee_id
           JOIN roles r ON r.id = a.role_id
           WHERE se.week_id = ? AND TRIM(se.label) NOT IN ('', 'OFF')
           ORDER BY e.name, se.day_index, a.role_id""",
        (int(week["id"]),),
    ).fetchall()
    warnings: list[dict[str, str]] = []
    by_employee_day: dict[tuple[int, int], list[sqlite3.Row]] = {}
    for row in rows:
        by_employee_day.setdefault((int(row["employee_id"]), int(row["day_index"])), []).append(row)

    school_days = school_days_for_week(week)
    weekly_minutes: dict[int, int] = {}
    incomplete_minor: set[int] = set()
    consent_late_nights: dict[int, int] = {}

    for (employee_id, day_index), day_rows in by_employee_day.items():
        employee = day_rows[0]
        work_day = week_start + timedelta(days=day_index)
        age = age_on(employee["date_of_birth"], work_day)
        all_intervals: list[tuple[int, int]] = []
        complete = True
        raw_interval_count = 0
        for row in day_rows:
            intervals, row_complete = shift_intervals(row["label"])
            all_intervals.extend(intervals)
            raw_interval_count += len(intervals)
            complete = complete and row_complete
            if row["label"] == "12:00 AM":
                warnings.append({"level": "warning", "text": f"{employee['name']} on {DAYS[day_index]} is entered as midnight; confirm this is intentional."})

        merged = merge_intervals(all_intervals)
        if len(merged) < raw_interval_count:
            warnings.append({"level": "error", "text": f"{employee['name']} has overlapping shifts or roles on {DAYS[day_index]}."})
        day_minutes = sum(end - start for start, end in merged)
        weekly_minutes[employee_id] = weekly_minutes.get(employee_id, 0) + day_minutes

        if age is None or age >= 18:
            continue
        if not complete:
            incomplete_minor.add(employee_id)
            warnings.append({"level": "error", "text": f"{employee['name']} is under 18 and has a shift without an explicit end time on {DAYS[day_index]}; compliance cannot be verified. Use a range such as 4:00 PM-9:00 PM."})
        for start, end in merged:
            if end - start >= 6 * 60:
                warnings.append({"level": "warning", "text": f"{employee['name']} is scheduled for at least 6 consecutive hours on {DAYS[day_index]}; Tennessee requires a 30-minute unpaid break that is not during or before the first hour."})

        school_start = parse_time_minutes(employee["school_start_time"] or "08:00") or 8 * 60
        school_end = parse_time_minutes(employee["school_end_time"] or "15:00") or 15 * 60
        if day_index in school_days and any(start < school_end and end > school_start for start, end in merged):
            warnings.append({"level": "error", "text": f"{employee['name']} is scheduled during entered school hours on {DAYS[day_index]}."})

        if age <= 15:
            max_day = 3 * 60 if day_index in school_days else 8 * 60
            if day_minutes > max_day:
                warnings.append({"level": "error", "text": f"{employee['name']} ({age}) is scheduled {day_minutes / 60:.1f} hours on {DAYS[day_index]}; the limit is {max_day / 60:.0f} hours for this day."})
            earliest = 7 * 60
            latest = 21 * 60 if not school_days and is_summer_evening(work_day) else 19 * 60
            if any(start < earliest or end > latest for start, end in merged):
                latest_text = time_to_12h(latest // 60)
                warnings.append({"level": "error", "text": f"{employee['name']} ({age}) is scheduled outside the permitted 7:00 AM-{latest_text} window on {DAYS[day_index]}."})
        else:
            next_day = work_day + timedelta(days=1)
            next_is_school_day = next_day.weekday() in school_days if next_day.weekday() < 7 else False
            if day_index == 6 and int(week["school_in_session"]):
                next_is_school_day = 0 in school_days
            if next_is_school_day and work_day.weekday() in {6, 0, 1, 2, 3}:
                latest_end = max((end for _, end in merged), default=0)
                has_consent = bool(employee["parental_late_consent"])
                if latest_end > (24 * 60 if has_consent else 22 * 60):
                    limit = "midnight with consent" if has_consent else "10:00 PM without a valid parental consent form"
                    warnings.append({"level": "error", "text": f"{employee['name']} ({age}) is scheduled too late on {DAYS[day_index]}; the limit is {limit} before a school day."})
                elif has_consent and latest_end > 22 * 60:
                    consent_late_nights[employee_id] = consent_late_nights.get(employee_id, 0) + 1

    employee_rows = {
        int(row["id"]): row for row in db.execute("SELECT id, name, date_of_birth FROM employees WHERE date_of_birth IS NOT NULL")
    }
    for employee_id, minutes in weekly_minutes.items():
        employee = employee_rows.get(employee_id)
        if not employee:
            continue
        age = age_on(employee["date_of_birth"], week_start)
        if age is not None and age <= 15 and employee_id not in incomplete_minor:
            max_week = 18 * 60 if school_days else 40 * 60
            if minutes > max_week:
                warnings.append({"level": "error", "text": f"{employee['name']} ({age}) is scheduled {minutes / 60:.1f} hours this week; the limit is {max_week / 60:.0f} hours."})
    for employee_id, nights in consent_late_nights.items():
        if nights > 3:
            employee = employee_rows.get(employee_id)
            warnings.append({"level": "error", "text": f"{employee['name']} is scheduled after 10:00 PM on {nights} school nights; parental consent permits no more than 3 such nights per week."})
    return warnings


def week_dates(week_start: date) -> list[date]:
    return [week_start + timedelta(days=index) for index in range(7)]


def surrounding_weeks(center: date, radius: int = 4):
    return [center + timedelta(days=7 * offset) for offset in range(-radius, radius + 1)]


def create_week_snapshot(db: sqlite3.Connection, week_id: int, reason: str) -> int:
    week = dict(db.execute("SELECT * FROM weeks WHERE id = ?", (week_id,)).fetchone())
    entries = [dict(row) for row in db.execute(
        """SELECT se.assignment_id, e.name, r.title AS role, r.subtitle, se.day_index, se.label
           FROM schedule_entries se
           JOIN employee_assignments a ON a.id = se.assignment_id
           JOIN employees e ON e.id = a.employee_id
           JOIN roles r ON r.id = a.role_id
           WHERE se.week_id = ? ORDER BY r.display_order, a.display_order, se.day_index""", (week_id,)
    )]
    cur = db.execute(
        "INSERT INTO week_versions (week_id, reason, snapshot_json) VALUES (?, ?, ?)",
        (week_id, reason, json.dumps({"week": week, "entries": entries}, indent=2)),
    )
    return int(cur.lastrowid)


def is_header_row(row) -> bool:
    return [str(row[index].value or "").strip().upper() for index in range(2, 9)] == DAYS


def import_history_from_workbook(path: Path) -> tuple[int, int, int]:
    book = load_workbook(path, data_only=True, read_only=True, keep_vba=True)
    rows_seen = 0
    shift_count = 0
    sheet_count = 0
    inserts = []
    for worksheet in book.worksheets:
        if worksheet.title.upper().startswith("SHEET"):
            continue
        sheet_count += 1
        current_role: tuple[str, str] | None = None
        sheet_rows = list(worksheet.iter_rows(min_row=1, max_row=130, min_col=1, max_col=9))
        for index, row in enumerate(sheet_rows):
            if is_header_row(row):
                title = str(row[0].value or "").strip().upper()
                subtitle = str(sheet_rows[index + 1][0].value or "").strip().upper() if index + 1 < len(sheet_rows) else ""
                current_role = (title, subtitle)
                continue
            if not current_role:
                continue
            name = str(row[0].value or "").strip().upper()
            if not name or name in DAY_WORDS or "SIGNATURE" in name or name.startswith("("):
                continue
            labels = [normalize_shift_label(cell.value, "12h") for cell in row[2:9]]
            if not any(label_is_usable(label, include_off=False) for label in labels):
                continue
            rows_seen += 1
            for day_index, label in enumerate(labels):
                if label_is_usable(label):
                    inserts.append((current_role[0], current_role[1], name, day_index, label, worksheet.title))
                    shift_count += 1
    book.close()
    with closing(get_db()) as db:
        db.execute("DELETE FROM shift_history")
        db.executemany(
            """INSERT INTO shift_history
               (role_title, role_subtitle, employee_name, day_index, label, source_sheet, count)
               VALUES (?, ?, ?, ?, ?, ?, 1)
               ON CONFLICT(role_title, role_subtitle, employee_name, day_index, label, source_sheet)
               DO UPDATE SET count = count + 1""", inserts
        )
        db.commit()
    return sheet_count, rows_seen, shift_count


def coverage_data(grouped, shifts):
    slots = list(range(9 * 60, 24 * 60, 60))
    tables = []
    incomplete = []
    for day_index in range(7):
        role_rows = []
        for role, employees_for_role in grouped:
            counts = []
            for slot in slots:
                count = 0
                for employee in employees_for_role:
                    label = shifts.get((employee.assignment_id, day_index), "")
                    intervals, complete = shift_intervals(label)
                    if label not in {"", "OFF"} and not complete:
                        incomplete.append(f"{employee.name} / {role.title} / {DAYS[day_index]}: {label}")
                    if any(start <= slot < end for start, end in intervals):
                        count += 1
                counts.append(count)
            role_rows.append((role, counts))
        tables.append(role_rows)
    return slots, tables, sorted(set(incomplete))


@app.before_request
def ensure_db() -> None:
    init_db()


@app.route("/")
def index():
    current = monday_for()
    with closing(get_db()) as db:
        week_rows = db.execute("SELECT * FROM weeks ORDER BY start_date DESC LIMIT 20").fetchall()
        history_rows = db.execute("SELECT COUNT(*) FROM shift_history").fetchone()[0]
        history_times = db.execute("SELECT COUNT(*) FROM shift_history WHERE label != 'OFF'").fetchone()[0]
    return render_template(
        "index.html", weeks=week_rows, current_week=current, week_nav=surrounding_weeks(current),
        seven_days=timedelta(days=6), history_rows=history_rows, history_times=history_times,
        data_path=str(DB_PATH), packaged=IS_PACKAGED,
    )


@app.route("/week", methods=["POST"])
def go_to_week():
    return redirect(url_for("edit_week", week_start=parse_week_start(request.form.get("week_start")).isoformat()))


@app.route("/week/<week_start>")
def edit_week(week_start: str):
    start = parse_week_start(week_start)
    with closing(get_db()) as db:
        week, grouped, shifts = grouped_schedule(db, start)
        suggestions = build_suggestions(db, grouped, start)
        warnings = collect_warnings(db, week, start)
        versions = db.execute("SELECT id, created_at, reason FROM week_versions WHERE week_id = ? ORDER BY id DESC", (week["id"],)).fetchall()
    return render_template(
        "schedule.html", week=week, week_start=start, week_end=start + timedelta(days=6),
        dates=week_dates(start), days=DAYS, grouped=grouped, shifts=shifts,
        suggestions=suggestions, warnings=warnings, versions=versions,
        school_days=school_days_for_week(week),
    )


@app.route("/api/week/<week_start>/shift", methods=["POST"])
def api_save_shift(week_start: str):
    data = request.get_json(force=True)
    start = parse_week_start(week_start)
    with closing(get_db()) as db:
        week = week_row(db, start)
        if week["published_at"]:
            return jsonify(ok=False, error="This schedule is published and locked."), 409
        employee = assignment_lookup(db).get(int(data["assignment_id"]))
        if not employee:
            return jsonify(ok=False, error="Assignment not found"), 404
        role = {role.id: role for role in roles(db)}[employee.role_id]
        label = save_shift_value(db, int(week["id"]), employee, role, int(data["day_index"]), data.get("label", ""))
        db.commit()
    return jsonify(ok=True, label=label)


@app.route("/api/week/<week_start>/fill-best", methods=["POST"])
def api_fill_best(week_start: str):
    start = parse_week_start(week_start)
    payload = {}
    with closing(get_db()) as db:
        week = week_row(db, start)
        if week["published_at"]:
            return jsonify(ok=False, error="This schedule is published and locked."), 409
        for role, employees_for_role in grouped_employees(db):
            for employee in employees_for_role:
                for day_index in range(7):
                    label = save_shift_value(db, int(week["id"]), employee, role, day_index, best_label_for_cell(db, role, employee, day_index))
                    payload[f"shift_{employee.assignment_id}_{day_index}"] = label
        db.commit()
    return jsonify(ok=True, shifts=payload)


@app.route("/api/week/<week_start>/fill-off", methods=["POST"])
def api_fill_off(week_start: str):
    start = parse_week_start(week_start)
    payload = {}
    with closing(get_db()) as db:
        week, grouped, shifts = grouped_schedule(db, start)
        if week["published_at"]:
            return jsonify(ok=False, error="This schedule is published and locked."), 409
        for role, employees_for_role in grouped:
            for employee in employees_for_role:
                for day_index in range(7):
                    if not shifts.get((employee.assignment_id, day_index), "").strip():
                        label = save_shift_value(db, int(week["id"]), employee, role, day_index, "OFF")
                        payload[f"shift_{employee.assignment_id}_{day_index}"] = label
        db.commit()
    return jsonify(ok=True, shifts=payload)


@app.route("/api/week/<week_start>/copy-previous", methods=["POST"])
def api_copy_previous(week_start: str):
    start = parse_week_start(week_start)
    payload = {}
    with closing(get_db()) as db:
        week = week_row(db, start)
        if week["published_at"]:
            return jsonify(ok=False, error="This schedule is published and locked."), 409
        previous = previous_week_options(db, start)
        if db.execute("SELECT COUNT(*) FROM schedule_entries WHERE week_id = ? AND label != ''", (week["id"],)).fetchone()[0]:
            create_week_snapshot(db, int(week["id"]), "Before replacing with previous week")
        for role, employees_for_role in grouped_employees(db):
            for employee in employees_for_role:
                for day_index in range(7):
                    label = save_shift_value(db, int(week["id"]), employee, role, day_index, previous.get((employee.assignment_id, day_index), ""))
                    payload[f"shift_{employee.assignment_id}_{day_index}"] = label
        db.commit()
    return jsonify(ok=True, shifts=payload)


@app.route("/week/<week_start>/save", methods=["POST"])
def save_week(week_start: str):
    start = parse_week_start(week_start)
    with closing(get_db()) as db:
        week = week_row(db, start)
        if week["published_at"]:
            flash("This schedule is published and locked. Reopen it before making corrections.")
            return redirect(url_for("edit_week", week_start=start.isoformat()))
        for role, employees_for_role in grouped_employees(db):
            for employee in employees_for_role:
                for day_index in range(7):
                    save_shift_value(db, int(week["id"]), employee, role, day_index, request.form.get(f"shift_{employee.assignment_id}_{day_index}", ""))
        db.commit()
    flash("Schedule saved.")
    return redirect(url_for("edit_week", week_start=start.isoformat()))


@app.route("/week/<week_start>/settings", methods=["POST"])
def save_week_settings(week_start: str):
    start = parse_week_start(week_start)
    with closing(get_db()) as db:
        week = week_row(db, start)
        if week["published_at"]:
            flash("Reopen the schedule before changing school-day settings.")
        else:
            school_in_session = 1 if request.form.get("school_in_session") else 0
            mask = "".join("1" if request.form.get(f"school_day_{index}") else "0" for index in range(7))
            db.execute("UPDATE weeks SET school_in_session = ?, school_day_mask = ? WHERE id = ?", (school_in_session, mask, week["id"]))
            db.commit()
            flash("School-day settings saved.")
    return redirect(url_for("edit_week", week_start=start.isoformat()))


@app.route("/week/<week_start>/publish", methods=["POST"])
def publish_week(week_start: str):
    start = parse_week_start(week_start)
    with closing(get_db()) as db:
        week = week_row(db, start)
        if not week["published_at"]:
            for role, employees_for_role in grouped_employees(db):
                for employee in employees_for_role:
                    for day_index in range(7):
                        field = f"shift_{employee.assignment_id}_{day_index}"
                        if field in request.form:
                            save_shift_value(db, int(week["id"]), employee, role, day_index, request.form[field])
            create_week_snapshot(db, int(week["id"]), "Published schedule")
            db.execute("UPDATE weeks SET published_at = CURRENT_TIMESTAMP WHERE id = ?", (week["id"],))
            db.commit()
            flash("Schedule published, snapshotted, and locked.")
    return redirect(url_for("edit_week", week_start=start.isoformat()))


@app.route("/week/<week_start>/reopen", methods=["POST"])
def reopen_week(week_start: str):
    start = parse_week_start(week_start)
    with closing(get_db()) as db:
        week = week_row(db, start)
        if week["published_at"]:
            create_week_snapshot(db, int(week["id"]), "Before reopening published schedule")
            db.execute("UPDATE weeks SET published_at = NULL WHERE id = ?", (week["id"],))
            db.commit()
            flash("Schedule reopened. The published copy remains in version history.")
    return redirect(url_for("edit_week", week_start=start.isoformat()))


@app.route("/week/<week_start>/copy-next", methods=["POST"])
def copy_next_week(week_start: str):
    start = parse_week_start(week_start)
    next_start = start + timedelta(days=7)
    with closing(get_db()) as db:
        source = week_row(db, start)
        target = week_row(db, next_start)
        existing = db.execute("SELECT COUNT(*) FROM schedule_entries WHERE week_id = ? AND label != ''", (target["id"],)).fetchone()[0]
        if existing:
            flash("The next week already contains schedule entries, so it was not overwritten.")
            return redirect(url_for("edit_week", week_start=next_start.isoformat()))
        if not source["published_at"]:
            for role, employees_for_role in grouped_employees(db):
                for employee in employees_for_role:
                    for day_index in range(7):
                        field = f"shift_{employee.assignment_id}_{day_index}"
                        if field in request.form:
                            save_shift_value(db, int(source["id"]), employee, role, day_index, request.form[field])
            create_week_snapshot(db, int(source["id"]), "Published when next week was created")
            db.execute("UPDATE weeks SET published_at = CURRENT_TIMESTAMP WHERE id = ?", (source["id"],))
        db.execute(
            """INSERT INTO schedule_entries (week_id, assignment_id, day_index, label)
               SELECT ?, assignment_id, day_index, label FROM schedule_entries WHERE week_id = ?""",
            (target["id"], source["id"]),
        )
        db.commit()
    flash("Current week published and preserved; next week created from it.")
    return redirect(url_for("edit_week", week_start=next_start.isoformat()))


@app.route("/week/<week_start>/versions")
def week_versions(week_start: str):
    start = parse_week_start(week_start)
    with closing(get_db()) as db:
        week = week_row(db, start)
        versions = db.execute("SELECT id, created_at, reason FROM week_versions WHERE week_id = ? ORDER BY id DESC", (week["id"],)).fetchall()
    return render_template("versions.html", week_start=start, versions=versions)


@app.route("/week/<week_start>/versions/<int:version_id>.json")
def download_week_version(week_start: str, version_id: int):
    start = parse_week_start(week_start)
    with closing(get_db()) as db:
        week = week_row(db, start)
        row = db.execute("SELECT snapshot_json FROM week_versions WHERE id = ? AND week_id = ?", (version_id, week["id"])).fetchone()
    if not row:
        return "Version not found", 404
    return Response(row["snapshot_json"], mimetype="application/json", headers={"Content-Disposition": f"attachment; filename=gaucho-schedule-{start.isoformat()}-version-{version_id}.json"})


@app.route("/coverage/<week_start>")
def coverage(week_start: str):
    start = parse_week_start(week_start)
    with closing(get_db()) as db:
        _, grouped, shifts = grouped_schedule(db, start)
    slots, tables, incomplete = coverage_data(grouped, shifts)
    return render_template("coverage.html", week_start=start, dates=week_dates(start), days=DAYS, slots=slots, tables=tables, incomplete=incomplete, format_minutes=format_minutes)


@app.route("/print/<week_start>")
def print_week(week_start: str):
    start = parse_week_start(week_start)
    with closing(get_db()) as db:
        _, grouped, shifts = grouped_schedule(db, start)
    return render_template("print.html", week_start=start, week_end=start + timedelta(days=6), dates=week_dates(start), days=DAYS, grouped=grouped, shifts=shifts)


@app.route("/preferences")
def preferences():
    return render_template("preferences.html")


@app.route("/import-history", methods=["GET", "POST"])
def import_history():
    if request.method == "POST":
        upload = request.files.get("schedule_workbook")
        if not upload or not upload.filename:
            flash("Choose the schedule workbook first.")
            return redirect(url_for("import_history"))
        suffix = Path(upload.filename).suffix or ".xlsx"
        with NamedTemporaryFile(delete=False, suffix=suffix) as temp:
            upload.save(temp.name)
            temp_path = Path(temp.name)
        try:
            sheet_count, row_count, shift_count = import_history_from_workbook(temp_path)
            flash(f"Imported {shift_count} shifts from {row_count} employee rows across {sheet_count} sheets.")
        finally:
            temp_path.unlink(missing_ok=True)
        return redirect(url_for("index"))
    return render_template("import_history.html")


@app.route("/employees")
def manage_employees():
    with closing(get_db()) as db:
        return render_template("employees.html", roles=roles(db), grouped=grouped_employees(db), archived_grouped=grouped_employees(db, archived=True))


@app.route("/api/employees/add", methods=["POST"])
def api_employee_add():
    data = request.get_json(force=True)
    name = (data.get("name") or "").strip().upper()
    role_id = int(data.get("role_id") or 1)
    if not name:
        return jsonify(ok=False, error="Name is required"), 400
    with closing(get_db()) as db:
        max_order = db.execute("SELECT COALESCE(MAX(display_order), 0) FROM employee_assignments WHERE role_id = ? AND archived_at IS NULL", (role_id,)).fetchone()[0]
        cur = db.execute("INSERT INTO employees (name, role_id, active, display_order) VALUES (?, ?, 1, ?)", (name, role_id, int(max_order) + 1))
        assignment = db.execute("INSERT INTO employee_assignments (employee_id, role_id, active, display_order) VALUES (?, ?, 1, ?)", (cur.lastrowid, role_id, int(max_order) + 1))
        db.commit()
    return jsonify(ok=True, id=cur.lastrowid, assignment_id=assignment.lastrowid)


@app.route("/api/employees/<int:employee_id>", methods=["PATCH"])
def api_employee_update(employee_id: int):
    data = request.get_json(force=True)
    fields = []
    params = []
    allowed = {
        "name": lambda value: str(value).strip().upper(),
        "date_of_birth": lambda value: str(value).strip() or None,
        "parental_late_consent": lambda value: 1 if value else 0,
        "school_start_time": lambda value: str(value).strip() or "08:00",
        "school_end_time": lambda value: str(value).strip() or "15:00",
    }
    for field, converter in allowed.items():
        if field in data:
            fields.append(f"{field} = ?")
            params.append(converter(data[field]))
    if not fields:
        return jsonify(ok=True)
    params.append(employee_id)
    with closing(get_db()) as db:
        db.execute(f"UPDATE employees SET {', '.join(fields)} WHERE id = ?", params)
        db.commit()
    return jsonify(ok=True)


@app.route("/api/employees/<int:employee_id>/roles", methods=["POST"])
def api_employee_add_role(employee_id: int):
    data = request.get_json(force=True)
    role_id = int(data["role_id"])
    with closing(get_db()) as db:
        existing = db.execute("SELECT id, archived_at FROM employee_assignments WHERE employee_id = ? AND role_id = ?", (employee_id, role_id)).fetchone()
        if existing:
            db.execute("UPDATE employee_assignments SET active = 1, archived_at = NULL WHERE id = ?", (existing["id"],))
            assignment_id = int(existing["id"])
        else:
            max_order = db.execute("SELECT COALESCE(MAX(display_order), 0) FROM employee_assignments WHERE role_id = ? AND archived_at IS NULL", (role_id,)).fetchone()[0]
            cur = db.execute("INSERT INTO employee_assignments (employee_id, role_id, active, display_order) VALUES (?, ?, 1, ?)", (employee_id, role_id, int(max_order) + 1))
            assignment_id = int(cur.lastrowid)
        db.execute("UPDATE employees SET active = 1, archived_at = NULL WHERE id = ?", (employee_id,))
        db.commit()
    return jsonify(ok=True, assignment_id=assignment_id)


@app.route("/api/assignments/<int:assignment_id>/archive", methods=["POST"])
def api_assignment_archive(assignment_id: int):
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    with closing(get_db()) as db:
        row = db.execute("SELECT employee_id FROM employee_assignments WHERE id = ?", (assignment_id,)).fetchone()
        if not row:
            return jsonify(ok=False, error="Assignment not found"), 404
        db.execute("UPDATE employee_assignments SET active = 0, archived_at = ? WHERE id = ?", (now, assignment_id))
        remaining = db.execute("SELECT COUNT(*) FROM employee_assignments WHERE employee_id = ? AND active = 1 AND archived_at IS NULL", (row["employee_id"],)).fetchone()[0]
        if not remaining:
            db.execute("UPDATE employees SET active = 0, archived_at = ? WHERE id = ?", (now, row["employee_id"]))
        db.commit()
    return jsonify(ok=True)


@app.route("/api/assignments/<int:assignment_id>/unarchive", methods=["POST"])
def api_assignment_unarchive(assignment_id: int):
    with closing(get_db()) as db:
        row = db.execute("SELECT employee_id FROM employee_assignments WHERE id = ?", (assignment_id,)).fetchone()
        if not row:
            return jsonify(ok=False, error="Assignment not found"), 404
        db.execute("UPDATE employee_assignments SET active = 1, archived_at = NULL WHERE id = ?", (assignment_id,))
        db.execute("UPDATE employees SET active = 1, archived_at = NULL WHERE id = ?", (row["employee_id"],))
        db.commit()
    return jsonify(ok=True)


@app.route("/api/employees/reorder", methods=["POST"])
def api_employee_reorder():
    data = request.get_json(force=True)
    with closing(get_db()) as db:
        for role_id, assignment_ids in (data.get("roles") or {}).items():
            for index, assignment_id in enumerate(assignment_ids, start=1):
                db.execute("UPDATE employee_assignments SET role_id = ?, display_order = ? WHERE id = ?", (int(role_id), index, int(assignment_id)))
        db.commit()
    return jsonify(ok=True)


@app.route("/backup")
def backup_database():
    init_db()
    with NamedTemporaryFile(delete=False, suffix=".sqlite3") as temp:
        temp_path = Path(temp.name)
    try:
        source = sqlite3.connect(DB_PATH)
        destination = sqlite3.connect(temp_path)
        source.backup(destination)
        destination.close()
        source.close()
        data = temp_path.read_bytes()
    finally:
        temp_path.unlink(missing_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d-%H%M")
    return send_file(io.BytesIO(data), as_attachment=True, download_name=f"gaucho-schedule-backup-{stamp}.sqlite3", mimetype="application/octet-stream")


@app.route("/export/<week_start>.csv")
def export_csv(week_start: str):
    start = parse_week_start(week_start)
    with closing(get_db()) as db:
        _, grouped, shifts = grouped_schedule(db, start)
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["GAUCHO URBANO EMPLOYEE SCHEDULE", start.isoformat(), "through", (start + timedelta(days=6)).isoformat()])
    for role, employees_for_role in grouped:
        writer.writerow([])
        writer.writerow([role.title, role.subtitle, *DAYS])
        writer.writerow(["", "", *[day.day for day in week_dates(start)]])
        for employee in employees_for_role:
            writer.writerow([employee.name, "", *[shifts.get((employee.assignment_id, index), "") for index in range(7)]])
    return Response(output.getvalue(), mimetype="text/csv", headers={"Content-Disposition": f"attachment; filename=gaucho-schedule-{start.isoformat()}.csv"})


@app.route("/export/<week_start>.xlsx")
def export_xlsx(week_start: str):
    start = parse_week_start(week_start)
    with closing(get_db()) as db:
        _, grouped, shifts = grouped_schedule(db, start)
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = start.strftime("%b %d")
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="EDEDED")
    row_number = 1
    date_format = "%-m/%-d/%Y" if os.name != "nt" else "%#m/%#d/%Y"
    worksheet.merge_cells(start_row=row_number, start_column=1, end_row=row_number, end_column=9)
    worksheet.cell(row=row_number, column=1, value=f"GAUCHO URBANO EMPLOYEE SCHEDULE    {start.strftime(date_format)} through {(start + timedelta(days=6)).strftime(date_format)}").font = Font(bold=True, size=14)
    row_number += 2
    for role, employees_for_role in grouped:
        worksheet.cell(row=row_number, column=1, value=f"{role.title} {role.subtitle}".strip()).font = Font(bold=True)
        worksheet.cell(row=row_number, column=1).fill = header_fill
        for column, day_name in enumerate(DAYS, start=3):
            cell = worksheet.cell(row=row_number, column=column, value=f"{day_name}\n{week_dates(start)[column - 3].day}")
            cell.font = Font(bold=True, italic=True)
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
            cell.fill = header_fill
        for column in range(1, 10):
            worksheet.cell(row=row_number, column=column).border = border
        row_number += 1
        for employee in employees_for_role:
            worksheet.cell(row=row_number, column=1, value=employee.name).font = Font(bold=True)
            for day_index in range(7):
                worksheet.cell(row=row_number, column=day_index + 3, value=shifts.get((employee.assignment_id, day_index), ""))
            for column in range(1, 10):
                worksheet.cell(row=row_number, column=column).border = border
                worksheet.cell(row=row_number, column=column).alignment = Alignment(horizontal="center", wrap_text=True)
            worksheet.cell(row=row_number, column=1).alignment = Alignment(horizontal="left")
            row_number += 1
        row_number += 1
    for column in range(1, 10):
        worksheet.column_dimensions[get_column_letter(column)].width = 14 if column > 1 else 18
    worksheet.page_setup.orientation = "portrait"
    worksheet.page_setup.fitToWidth = 1
    worksheet.page_setup.fitToHeight = 2
    file_stream = io.BytesIO()
    workbook.save(file_stream)
    file_stream.seek(0)
    return send_file(file_stream, as_attachment=True, download_name=f"gaucho-schedule-{start.isoformat()}.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def open_browser() -> None:
    port = int(os.environ.get("PORT", "5000"))
    webbrowser.open(f"http://127.0.0.1:{port}")


if __name__ == "__main__":
    init_db()
    default_host = "0.0.0.0" if os.environ.get("REPL_ID") else "127.0.0.1"
    if IS_PACKAGED and default_host == "127.0.0.1":
        threading.Timer(1.0, open_browser).start()
    debug = os.environ.get("GAUCHO_SCHEDULE_DEBUG", "0") == "1" and not IS_PACKAGED
    app.run(host=os.environ.get("GAUCHO_SCHEDULE_HOST", default_host), port=int(os.environ.get("PORT", "5000")), debug=debug, use_reloader=debug)
