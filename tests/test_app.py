import tempfile
import unittest
from datetime import date
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

    def test_split_shift_normalization_and_hours(self):
        label = schedule_app.normalize_shift_label("9-12 / 4-9")
        self.assertEqual(label, "9:00 AM-12:00 PM / 4:00 PM-9:00 PM")
        intervals, complete = schedule_app.shift_intervals(label)
        self.assertTrue(complete)
        self.assertEqual(sum(end - start for start, end in intervals), 8 * 60)

    def test_incomplete_shift_is_detected(self):
        intervals, complete = schedule_app.shift_intervals("4:00 PM")
        self.assertEqual(intervals, [])
        self.assertFalse(complete)

    def test_minor_late_shift_warning(self):
        week_start = date(2026, 8, 24)
        with schedule_app.closing(schedule_app.get_db()) as db:
            employee = db.execute("SELECT * FROM employees LIMIT 1").fetchone()
            assignment = db.execute("SELECT * FROM employee_assignments WHERE employee_id = ? LIMIT 1", (employee["id"],)).fetchone()
            db.execute("UPDATE employees SET date_of_birth = '2010-08-26' WHERE id = ?", (employee["id"],))
            week_id = schedule_app.get_or_create_week(db, week_start)
            db.execute("INSERT INTO schedule_entries (week_id, assignment_id, day_index, label) VALUES (?, ?, 0, ?)", (week_id, assignment["id"], "4:00 PM-11:00 PM"))
            db.commit()
            week = db.execute("SELECT * FROM weeks WHERE id = ?", (week_id,)).fetchone()
            warnings = schedule_app.collect_warnings(db, week, week_start)
        self.assertTrue(any("outside the permitted" in warning["text"] for warning in warnings))

    def test_same_employee_can_have_multiple_roles(self):
        with schedule_app.closing(schedule_app.get_db()) as db:
            employee = db.execute("SELECT * FROM employees LIMIT 1").fetchone()
            existing_role = employee["role_id"]
            other_role = db.execute("SELECT id FROM roles WHERE id != ? LIMIT 1", (existing_role,)).fetchone()[0]
            db.execute("INSERT INTO employee_assignments (employee_id, role_id, active, display_order) VALUES (?, ?, 1, 99)", (employee["id"], other_role))
            db.commit()
            count = db.execute("SELECT COUNT(*) FROM employee_assignments WHERE employee_id = ?", (employee["id"],)).fetchone()[0]
        self.assertEqual(count, 2)


if __name__ == "__main__":
    unittest.main()
