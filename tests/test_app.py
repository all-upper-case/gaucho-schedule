import tempfile
import unittest
from pathlib import Path

import app as schedule_app


class GauchoScheduleTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.original_db = schedule_app.DB_PATH
        schedule_app.DB_PATH = Path(self.temp_dir.name) / "test.sqlite3"
        schedule_app.init_db()

    def tearDown(self):
        schedule_app.DB_PATH = self.original_db
        self.temp_dir.cleanup()

    def test_two_report_times_normalize_without_end_times(self):
        label = schedule_app.normalize_shift_label("9 / 4")
        self.assertEqual(label, "9:00 AM / 4:00 PM")

    def test_pasted_range_keeps_only_start_time(self):
        label = schedule_app.normalize_shift_label("4:00 PM-9:00 PM")
        self.assertEqual(label, "4:00 PM")

    def test_same_employee_can_have_multiple_roles(self):
        with schedule_app.closing(schedule_app.get_db()) as db:
            employee = db.execute("SELECT * FROM employees LIMIT 1").fetchone()
            existing_role = employee["role_id"]
            other_role = db.execute("SELECT id FROM roles WHERE id != ? LIMIT 1", (existing_role,)).fetchone()[0]
            db.execute("INSERT INTO employee_assignments (employee_id, role_id, active, display_order) VALUES (?, ?, 1, 99)", (employee["id"], other_role))
            db.commit()
            count = db.execute("SELECT COUNT(*) FROM employee_assignments WHERE employee_id = ?", (employee["id"],)).fetchone()[0]
        self.assertEqual(count, 2)

    def test_headcount_counts_split_report_time_once(self):
        with schedule_app.closing(schedule_app.get_db()) as db:
            week_id = schedule_app.get_or_create_week(db, schedule_app.monday_for())
            assignment = db.execute("SELECT id FROM employee_assignments LIMIT 1").fetchone()[0]
            db.execute("INSERT INTO schedule_entries (week_id, assignment_id, day_index, label) VALUES (?, ?, 0, '9:00 AM / 4:00 PM')", (week_id, assignment))
            db.commit()
            grouped = schedule_app.grouped_employees(db)
            shifts = schedule_app.shift_map(db, week_id)
        rows = schedule_app.coverage_data(grouped, shifts)
        matching = next(row for row in rows if row[0].id == 1)
        self.assertEqual(matching[1][0], 1)


if __name__ == "__main__":
    unittest.main()
