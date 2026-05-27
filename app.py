from __future__ import annotations

import csv
import io
import os
import re
import sqlite3
from contextlib import closing
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from pathlib import Path
from tempfile import NamedTemporaryFile

from flask import Flask, Response, flash, jsonify, redirect, render_template, request, send_file, url_for
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)
DB_PATH = Path(os.environ.get("GAUCHO_SCHEDULE_DB", DATA_DIR / "gaucho_schedule.sqlite3"))

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-local-only-change-before-hosting")

DAYS = ["MON", "TUES", "WEDS", "THURS", "FRI", "SAT", "SUN"]
DAY_WORDS = set(DAYS + ["THU"])
GLOBAL_SHIFT_OPTIONS = ["3:00 PM", "4:00 PM", "5:00 PM", "2:30 PM", "10:00 AM", "10:30 AM", "11:00 AM", "12:00 PM", "1:00 PM", "2:00 PM", "CLOSE", "OFF"]

DEFAULT_ROLES = [(1, "SERVERS", "(ALCOHOL)"), (2, "SERVERS", "(FOOD ONLY)"), (3, "GAUCHOS", "(MEAT)"), (4, "SERVER", "ASSISTANT"), (5, "HOST", ""), (6, "KITCHEN", ""), (7, "DISHWASHER", "")]
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
class Employee:
    id: int
    name: str
    role_id: int
    active: int
    display_order: int
    archived_at: str | None = None


def get_db() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def ensure_column(db: sqlite3.Connection, table: str, column: str, ddl: str) -> None:
    cols = {row["name"] for row in db.execute(f"PRAGMA table_info({table})")}
    if column not in cols:
        db.execute(f"ALTER TABLE {table} ADD COLUMN {column} {ddl}")


def init_db() -> None:
    with closing(get_db()) as db:
        db.executescript(
            """
            CREATE TABLE IF NOT EXISTS roles (id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT NOT NULL, subtitle TEXT NOT NULL DEFAULT '', display_order INTEGER NOT NULL DEFAULT 0);
            CREATE TABLE IF NOT EXISTS employees (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, role_id INTEGER NOT NULL REFERENCES roles(id), active INTEGER NOT NULL DEFAULT 1, display_order INTEGER NOT NULL DEFAULT 0);
            CREATE TABLE IF NOT EXISTS weeks (id INTEGER PRIMARY KEY AUTOINCREMENT, start_date TEXT NOT NULL UNIQUE, notes TEXT NOT NULL DEFAULT '', created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP);
            CREATE TABLE IF NOT EXISTS shifts (id INTEGER PRIMARY KEY AUTOINCREMENT, week_id INTEGER NOT NULL REFERENCES weeks(id) ON DELETE CASCADE, employee_id INTEGER NOT NULL REFERENCES employees(id) ON DELETE CASCADE, day_index INTEGER NOT NULL CHECK(day_index BETWEEN 0 AND 6), label TEXT NOT NULL DEFAULT 'OFF', UNIQUE(week_id, employee_id, day_index));
            CREATE TABLE IF NOT EXISTS shift_history (id INTEGER PRIMARY KEY AUTOINCREMENT, role_title TEXT NOT NULL, role_subtitle TEXT NOT NULL DEFAULT '', employee_name TEXT NOT NULL, day_index INTEGER NOT NULL CHECK(day_index BETWEEN 0 AND 6), label TEXT NOT NULL, source_sheet TEXT NOT NULL DEFAULT '', count INTEGER NOT NULL DEFAULT 1, UNIQUE(role_title, role_subtitle, employee_name, day_index, label, source_sheet));
            """
        )
        ensure_column(db, "employees", "archived_at", "TEXT")
        if db.execute("SELECT COUNT(*) FROM roles").fetchone()[0] == 0:
            db.executemany("INSERT INTO roles (display_order, title, subtitle) VALUES (?, ?, ?)", [(o, t, s) for o, t, s in DEFAULT_ROLES])
        if db.execute("SELECT COUNT(*) FROM employees").fetchone()[0] == 0:
            counters: dict[int, int] = {}
            rows = []
            for name, role_id in DEFAULT_EMPLOYEES:
                counters[role_id] = counters.get(role_id, 0) + 1
                rows.append((name, role_id, 1, counters[role_id]))
            db.executemany("INSERT INTO employees (name, role_id, active, display_order) VALUES (?, ?, ?, ?)", rows)
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


def roles(db: sqlite3.Connection) -> list[Role]:
    return [Role(**dict(r)) for r in db.execute("SELECT * FROM roles ORDER BY display_order, id")]


def role_by_id(db: sqlite3.Connection) -> dict[int, Role]:
    return {role.id: role for role in roles(db)}


def employees(db: sqlite3.Connection, include_inactive: bool = False, archived: bool = False) -> list[Employee]:
    if archived:
        where = "WHERE archived_at IS NOT NULL"
    elif include_inactive:
        where = ""
    else:
        where = "WHERE active = 1 AND archived_at IS NULL"
    return [Employee(**dict(r)) for r in db.execute(f"SELECT * FROM employees {where} ORDER BY role_id, display_order, name, id")]


def grouped_employees(db: sqlite3.Connection, archived: bool = False):
    all_roles = roles(db)
    by_role: dict[int, list[Employee]] = {role.id: [] for role in all_roles}
    for emp in employees(db, archived=archived):
        by_role.setdefault(emp.role_id, []).append(emp)
    return [(role, by_role.get(role.id, [])) for role in all_roles]


def shift_map(db: sqlite3.Connection, week_id: int) -> dict[tuple[int, int], str]:
    rows = db.execute("SELECT employee_id, day_index, label FROM shifts WHERE week_id = ?", (week_id,)).fetchall()
    return {(int(r["employee_id"]), int(r["day_index"])): r["label"] for r in rows}


def grouped_schedule(db: sqlite3.Connection, week_start: date):
    week_id = get_or_create_week(db, week_start)
    return week_id, grouped_employees(db), shift_map(db, week_id)


def time_to_12h(hour: int, minute: int = 0) -> str:
    hour %= 24
    return f"{hour % 12 or 12}:{minute:02d} {'AM' if hour < 12 else 'PM'}"


def format_minutes(total_minutes: int, output_format: str = "12h") -> str:
    total_minutes %= 24 * 60
    hour = total_minutes // 60
    minute = total_minutes % 60
    return f"{hour:02d}:{minute:02d}" if output_format == "24h" else time_to_12h(hour, minute)


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
        if float(value).is_integer():
            return normalize_shift_label(str(int(value)), output_format)
        return str(value).strip()
    raw = str(value).strip()
    if not raw:
        return ""
    upper = raw.upper()
    if upper in {"0", "OFF", "OOF"}:
        return "OFF"
    if upper in {"CL", "CLOSE", "CLOSING"}:
        return "CLOSE"
    if any(separator in raw for separator in [";", "-", "–", "—"]):
        pieces = [p.strip() for p in re.split(r"[;\-–—]+", raw) if p.strip()]
        if len(pieces) == 2:
            first = normalize_shift_label(pieces[0], output_format)
            second = normalize_shift_label(pieces[1], output_format)
            if is_real_time_label(first) and is_real_time_label(second):
                return f"{first}-{second}"
    compact = re.fullmatch(r"\d{3,4}", raw)
    if compact:
        hour = int(raw[:-2]); minute = int(raw[-2:])
        if hour <= 23 and minute < 60:
            if hour <= 7: hour += 12
            return format_minutes(hour * 60 + minute, output_format)
    match = re.fullmatch(r"(\d{1,2})(?::(\d{1,2}))?\s*([AaPp][Mm])?", raw.replace(";", ":").replace(".", ":"))
    if match:
        hour = int(match.group(1)); minute = int(match.group(2) or 0); suffix = match.group(3).upper() if match.group(3) else None
        if minute >= 60: return upper
        if suffix == "PM" and hour < 12: hour += 12
        if suffix == "AM" and hour == 12: hour = 0
        if not suffix and hour <= 7: hour += 12
        if hour <= 23: return format_minutes(hour * 60 + minute, output_format)
    return upper


def is_real_time_label(label: str) -> bool:
    return bool(re.search(r"\d{1,2}:\d{2}\s*(AM|PM)?", label))


def label_is_usable(label: str, include_off: bool = True) -> bool:
    if not label: return False
    if label == "OFF": return include_off
    if label == "CLOSE": return True
    if label.upper() in DAY_WORDS or label.upper() == "THROUGH": return False
    return is_real_time_label(label)


def unique_options(*groups: list[str]) -> list[str]:
    seen: set[str] = set(); output: list[str] = []
    for group in groups:
        for item in group:
            clean = (item or "").strip()
            if clean and clean not in seen:
                seen.add(clean); output.append(clean)
    return output


def previous_week_options(db: sqlite3.Connection, week_start: date) -> dict[tuple[int, int], str]:
    row = db.execute("SELECT id FROM weeks WHERE start_date = ?", ((week_start - timedelta(days=7)).isoformat(),)).fetchone()
    return shift_map(db, int(row["id"])) if row else {}


def history_options_for_cell(db: sqlite3.Connection, role: Role, employee: Employee, day_index: int) -> list[str]:
    rows = db.execute("""SELECT label, SUM(count) AS total FROM shift_history WHERE role_title = ? AND role_subtitle = ? AND employee_name = ? AND day_index = ? AND label != 'OFF' GROUP BY label ORDER BY total DESC, label LIMIT 12""", (role.title, role.subtitle, employee.name.upper(), day_index)).fetchall()
    return [r["label"] for r in rows]


def role_options_for_cell(db: sqlite3.Connection, role: Role, day_index: int) -> list[str]:
    rows = db.execute("""SELECT label, SUM(count) AS total FROM shift_history WHERE role_title = ? AND role_subtitle = ? AND day_index = ? AND label != 'OFF' GROUP BY label ORDER BY total DESC, label LIMIT 10""", (role.title, role.subtitle, day_index)).fetchall()
    return [r["label"] for r in rows]


def global_options_for_cell(db: sqlite3.Connection, day_index: int) -> list[str]:
    rows = db.execute("""SELECT label, SUM(count) AS total FROM shift_history WHERE day_index = ? AND label != 'OFF' GROUP BY label ORDER BY total DESC, label LIMIT 10""", (day_index,)).fetchall()
    return [r["label"] for r in rows]


def best_label_for_cell(db: sqlite3.Connection, role: Role, emp: Employee, day_index: int) -> str:
    for source in [history_options_for_cell(db, role, emp, day_index), role_options_for_cell(db, role, day_index), global_options_for_cell(db, day_index), GLOBAL_SHIFT_OPTIONS]:
        for label in source:
            if label and label != "OFF":
                return label
    return "OFF"


def build_suggestions(db: sqlite3.Connection, grouped, week_start: date) -> dict[tuple[int, int], list[str]]:
    previous = previous_week_options(db, week_start)
    suggestions: dict[tuple[int, int], list[str]] = {}
    for role, emps in grouped:
        for emp in emps:
            for day_index in range(7):
                prev = previous.get((emp.id, day_index))
                suggestions[(emp.id, day_index)] = unique_options([prev] if prev else [], history_options_for_cell(db, role, emp, day_index), role_options_for_cell(db, role, day_index), global_options_for_cell(db, day_index), GLOBAL_SHIFT_OPTIONS)[:20]
    return suggestions


def save_shift_value(db: sqlite3.Connection, week_id: int, emp: Employee, role: Role, day_index: int, raw_label: str) -> str:
    label = normalize_shift_label(raw_label or "OFF", "12h") or "OFF"
    db.execute("""INSERT INTO shifts (week_id, employee_id, day_index, label) VALUES (?, ?, ?, ?) ON CONFLICT(week_id, employee_id, day_index) DO UPDATE SET label = excluded.label""", (week_id, emp.id, day_index, label))
    if label_is_usable(label):
        db.execute("""INSERT INTO shift_history (role_title, role_subtitle, employee_name, day_index, label, source_sheet, count) VALUES (?, ?, ?, ?, ?, 'manual-app-entry', 1) ON CONFLICT(role_title, role_subtitle, employee_name, day_index, label, source_sheet) DO UPDATE SET count = count + 1""", (role.title, role.subtitle, emp.name.upper(), day_index, label))
    return label


def collect_warnings(db: sqlite3.Connection, week_id: int) -> list[str]:
    rows = db.execute("""SELECT employees.name, shifts.day_index, shifts.label FROM shifts JOIN employees ON employees.id = shifts.employee_id WHERE shifts.week_id = ? AND TRIM(shifts.label) != '' AND UPPER(TRIM(shifts.label)) != 'OFF' ORDER BY employees.name, shifts.day_index""", (week_id,)).fetchall()
    return [f"{r['name']} on {DAYS[int(r['day_index'])]}: {r['label']} is midnight; confirm this is intentional." for r in rows if r["label"] == "12:00 AM"]


def week_dates(week_start: date) -> list[date]:
    return [week_start + timedelta(days=i) for i in range(7)]


def surrounding_weeks(center: date, radius: int = 4):
    return [center + timedelta(days=7 * offset) for offset in range(-radius, radius + 1)]


def is_header_row(row) -> bool:
    return [str(row[i].value or "").strip().upper() for i in range(2, 9)] == DAYS


def import_history_from_workbook(path: Path) -> tuple[int, int, int]:
    book = load_workbook(path, data_only=True, read_only=True, keep_vba=True)
    rows_seen = 0; shift_count = 0; sheet_count = 0; inserts = []
    for ws in book.worksheets:
        if ws.title.upper().startswith("SHEET"): continue
        sheet_count += 1; current_role: tuple[str, str] | None = None
        sheet_rows = list(ws.iter_rows(min_row=1, max_row=130, min_col=1, max_col=9))
        for idx, row in enumerate(sheet_rows):
            if is_header_row(row):
                title = str(row[0].value or "").strip().upper()
                subtitle = str(sheet_rows[idx + 1][0].value or "").strip().upper() if idx + 1 < len(sheet_rows) else ""
                current_role = (title, subtitle); continue
            if not current_role: continue
            name = str(row[0].value or "").strip().upper()
            if not name or name in DAY_WORDS or "SIGNATURE" in name or name.startswith("("): continue
            row_labels = [normalize_shift_label(cell.value, "12h") for cell in row[2:9]]
            if not any(label_is_usable(label, include_off=False) for label in row_labels): continue
            rows_seen += 1
            for day_index, label in enumerate(row_labels):
                if label_is_usable(label):
                    inserts.append((current_role[0], current_role[1], name, day_index, label, ws.title)); shift_count += 1
    book.close()
    with closing(get_db()) as db:
        db.execute("DELETE FROM shift_history")
        db.executemany("""INSERT INTO shift_history (role_title, role_subtitle, employee_name, day_index, label, source_sheet, count) VALUES (?, ?, ?, ?, ?, ?, 1) ON CONFLICT(role_title, role_subtitle, employee_name, day_index, label, source_sheet) DO UPDATE SET count = count + 1""", inserts)
        db.commit()
    return sheet_count, rows_seen, shift_count

@app.before_request
def ensure_db() -> None:
    init_db()

@app.route("/")
def index():
    current = monday_for()
    with closing(get_db()) as db:
        week_rows = db.execute("SELECT * FROM weeks ORDER BY start_date DESC LIMIT 12").fetchall()
        history_rows = db.execute("SELECT COUNT(*) FROM shift_history").fetchone()[0]
        history_times = db.execute("SELECT COUNT(*) FROM shift_history WHERE label != 'OFF'").fetchone()[0]
    return render_template("index.html", weeks=week_rows, current_week=current, week_nav=surrounding_weeks(current), history_rows=history_rows, history_times=history_times)

@app.route("/week", methods=["POST"])
def go_to_week():
    return redirect(url_for("edit_week", week_start=parse_week_start(request.form.get("week_start")).isoformat()))

@app.route("/week/<week_start>")
def edit_week(week_start: str):
    start = parse_week_start(week_start)
    with closing(get_db()) as db:
        week_id, grouped, shifts = grouped_schedule(db, start)
        suggestions = build_suggestions(db, grouped, start)
        warnings = collect_warnings(db, week_id)
    return render_template("schedule.html", week_start=start, week_end=start + timedelta(days=6), dates=week_dates(start), days=DAYS, grouped=grouped, shifts=shifts, suggestions=suggestions, warnings=warnings)

@app.route("/api/week/<week_start>/shift", methods=["POST"])
def api_save_shift(week_start: str):
    data = request.get_json(force=True)
    start = parse_week_start(week_start)
    with closing(get_db()) as db:
        emp = Employee(**dict(db.execute("SELECT * FROM employees WHERE id = ?", (int(data["employee_id"]),)).fetchone()))
        role = role_by_id(db)[emp.role_id]
        week_id = get_or_create_week(db, start)
        label = save_shift_value(db, week_id, emp, role, int(data["day_index"]), data.get("label", "OFF"))
        db.commit()
    return jsonify(ok=True, label=label)

@app.route("/api/week/<week_start>/fill-best", methods=["POST"])
def api_fill_best(week_start: str):
    start = parse_week_start(week_start)
    payload = {}
    with closing(get_db()) as db:
        week_id = get_or_create_week(db, start)
        for role, emps in grouped_employees(db):
            for emp in emps:
                for day_index in range(7):
                    label = save_shift_value(db, week_id, emp, role, day_index, best_label_for_cell(db, role, emp, day_index))
                    payload[f"shift_{emp.id}_{day_index}"] = label
        db.commit()
    return jsonify(ok=True, shifts=payload)

@app.route("/api/week/<week_start>/copy-previous", methods=["POST"])
def api_copy_previous(week_start: str):
    start = parse_week_start(week_start)
    payload = {}
    with closing(get_db()) as db:
        previous = previous_week_options(db, start)
        week_id = get_or_create_week(db, start)
        role_lookup = role_by_id(db)
        for emp in employees(db):
            role = role_lookup[emp.role_id]
            for day_index in range(7):
                label = previous.get((emp.id, day_index), "OFF")
                label = save_shift_value(db, week_id, emp, role, day_index, label)
                payload[f"shift_{emp.id}_{day_index}"] = label
        db.commit()
    return jsonify(ok=True, shifts=payload)

@app.route("/week/<week_start>/save", methods=["POST"])
def save_week(week_start: str):
    start = parse_week_start(week_start)
    with closing(get_db()) as db:
        week_id = get_or_create_week(db, start); role_lookup = role_by_id(db)
        for emp in employees(db):
            role = role_lookup[emp.role_id]
            for day_index in range(7):
                save_shift_value(db, week_id, emp, role, day_index, request.form.get(f"shift_{emp.id}_{day_index}", "OFF"))
        db.commit()
    flash("Schedule saved.")
    return redirect(url_for("edit_week", week_start=start.isoformat()))

@app.route("/week/<week_start>/copy-next", methods=["POST"])
def copy_next_week(week_start: str):
    start = parse_week_start(week_start); next_start = start + timedelta(days=7)
    with closing(get_db()) as db:
        source_id = get_or_create_week(db, start); target_id = get_or_create_week(db, next_start)
        db.execute("DELETE FROM shifts WHERE week_id = ?", (target_id,))
        db.execute("INSERT INTO shifts (week_id, employee_id, day_index, label) SELECT ?, employee_id, day_index, label FROM shifts WHERE week_id = ?", (target_id, source_id))
        db.commit()
    flash("Next week created by copying the current week.")
    return redirect(url_for("edit_week", week_start=next_start.isoformat()))

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
            flash("Choose the schedule workbook first."); return redirect(url_for("import_history"))
        suffix = Path(upload.filename).suffix or ".xlsx"
        with NamedTemporaryFile(delete=False, suffix=suffix) as temp:
            upload.save(temp.name); temp_path = Path(temp.name)
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
    name = (data.get("name") or "").strip().upper(); role_id = int(data.get("role_id") or 1)
    if not name: return jsonify(ok=False, error="Name is required"), 400
    with closing(get_db()) as db:
        max_order = db.execute("SELECT COALESCE(MAX(display_order), 0) FROM employees WHERE role_id = ? AND archived_at IS NULL", (role_id,)).fetchone()[0]
        cur = db.execute("INSERT INTO employees (name, role_id, active, display_order, archived_at) VALUES (?, ?, 1, ?, NULL)", (name, role_id, int(max_order) + 1))
        db.commit()
    return jsonify(ok=True, id=cur.lastrowid)

@app.route("/api/employees/<int:employee_id>", methods=["PATCH"])
def api_employee_update(employee_id: int):
    data = request.get_json(force=True)
    fields = [] ; params = []
    if "name" in data: fields.append("name = ?"); params.append(str(data["name"]).strip().upper())
    if "role_id" in data: fields.append("role_id = ?"); params.append(int(data["role_id"]))
    if "display_order" in data: fields.append("display_order = ?"); params.append(int(data["display_order"]))
    if not fields: return jsonify(ok=True)
    params.append(employee_id)
    with closing(get_db()) as db:
        db.execute(f"UPDATE employees SET {', '.join(fields)} WHERE id = ?", params); db.commit()
    return jsonify(ok=True)

@app.route("/api/employees/<int:employee_id>/archive", methods=["POST"])
def api_employee_archive(employee_id: int):
    with closing(get_db()) as db:
        db.execute("UPDATE employees SET active = 0, archived_at = ? WHERE id = ?", (datetime.now().strftime("%Y-%m-%d %H:%M"), employee_id)); db.commit()
    return jsonify(ok=True)

@app.route("/api/employees/<int:employee_id>/unarchive", methods=["POST"])
def api_employee_unarchive(employee_id: int):
    with closing(get_db()) as db:
        db.execute("UPDATE employees SET active = 1, archived_at = NULL WHERE id = ?", (employee_id,)); db.commit()
    return jsonify(ok=True)

@app.route("/api/employees/reorder", methods=["POST"])
def api_employee_reorder():
    data = request.get_json(force=True)
    with closing(get_db()) as db:
        for role_id, ids in (data.get("roles") or {}).items():
            for index, employee_id in enumerate(ids, start=1):
                db.execute("UPDATE employees SET role_id = ?, display_order = ? WHERE id = ?", (int(role_id), index, int(employee_id)))
        db.commit()
    return jsonify(ok=True)

@app.route("/export/<week_start>.csv")
def export_csv(week_start: str):
    start = parse_week_start(week_start)
    with closing(get_db()) as db:
        _, grouped, shifts = grouped_schedule(db, start)
    output = io.StringIO(); writer = csv.writer(output)
    writer.writerow(["GAUCHO URBANO EMPLOYEE SCHEDULE", start.isoformat(), "through", (start + timedelta(days=6)).isoformat()])
    for role, emps in grouped:
        writer.writerow([]); writer.writerow([role.title, role.subtitle, *DAYS]); writer.writerow(["", "", *[d.day for d in week_dates(start)]])
        for emp in emps: writer.writerow([emp.name, "", *[shifts.get((emp.id, i), "OFF") for i in range(7)]])
    return Response(output.getvalue(), mimetype="text/csv", headers={"Content-Disposition": f"attachment; filename=gaucho-schedule-{start.isoformat()}.csv"})

@app.route("/export/<week_start>.xlsx")
def export_xlsx(week_start: str):
    start = parse_week_start(week_start)
    with closing(get_db()) as db:
        _, grouped, shifts = grouped_schedule(db, start)
    wb = Workbook(); ws = wb.active; ws.title = start.strftime("%b %d")
    thin = Side(style="thin"); border = Border(left=thin, right=thin, top=thin, bottom=thin); header_fill = PatternFill("solid", fgColor="EDEDED")
    row = 1; date_fmt = "%-m/%-d/%Y" if os.name != "nt" else "%#m/%#d/%Y"
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=9)
    ws.cell(row=row, column=1, value=f"GAUCHO URBANO EMPLOYEE SCHEDULE    {start.strftime(date_fmt)} through {(start + timedelta(days=6)).strftime(date_fmt)}").font = Font(bold=True, size=14)
    row += 2
    for role, emps in grouped:
        ws.cell(row=row, column=1, value=f"{role.title} {role.subtitle}".strip()).font = Font(bold=True)
        ws.cell(row=row, column=1).fill = header_fill
        for idx, day in enumerate(DAYS, start=3):
            c = ws.cell(row=row, column=idx, value=f"{day}\n{week_dates(start)[idx - 3].day}"); c.font = Font(bold=True, italic=True); c.alignment = Alignment(horizontal="center", wrap_text=True); c.fill = header_fill
        for col in range(1, 10): ws.cell(row=row, column=col).border = border
        row += 1
        for emp in emps:
            ws.cell(row=row, column=1, value=emp.name).font = Font(bold=True)
            for day_index in range(7): ws.cell(row=row, column=day_index + 3, value=shifts.get((emp.id, day_index), "OFF"))
            for col in range(1, 10):
                ws.cell(row=row, column=col).border = border; ws.cell(row=row, column=col).alignment = Alignment(horizontal="center")
            ws.cell(row=row, column=1).alignment = Alignment(horizontal="left")
            row += 1
        row += 1
    for col in range(1, 10): ws.column_dimensions[get_column_letter(col)].width = 14 if col > 1 else 18
    ws.page_setup.orientation = "portrait"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 2
    file_stream = io.BytesIO(); wb.save(file_stream); file_stream.seek(0)
    return send_file(file_stream, as_attachment=True, download_name=f"gaucho-schedule-{start.isoformat()}.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    init_db()
    default_host = "0.0.0.0" if os.environ.get("REPL_ID") else "127.0.0.1"
    app.run(host=os.environ.get("GAUCHO_SCHEDULE_HOST", default_host), port=int(os.environ.get("PORT", "5000")), debug=True)
