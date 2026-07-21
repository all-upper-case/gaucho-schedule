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

    def test_pos_classifier_ignores_meal_reclock_and_keeps_split_start(self):
        rows = [
            {"Employee": "Smith, Alex", "Job": "Kitchen Aid", "_date": date(2026, 7, 1), "_in": 9 * 60, "_out": 12 * 60, "_line": 2},
            {"Employee": "Smith, Alex", "Job": "Kitchen Aid", "_date": date(2026, 7, 1), "_in": 12 * 60 + 30, "_out": 14 * 60, "_line": 3},
            {"Employee": "Smith, Alex", "Job": "Kitchen Aid", "_date": date(2026, 7, 1), "_in": 16 * 60, "_out": 21 * 60, "_line": 4},
        ]
        patterns, stats = schedule_app.classify_pos_starts(rows)
        self.assertEqual([pattern["label"] for pattern in patterns], ["9:00 AM / 4:00 PM"])
        self.assertEqual(stats["meal_or_reclock"], 1)
        self.assertEqual(stats["split_start"], 1)

    def test_pos_name_aliases_match_printed_schedule_style(self):
        aliases = schedule_app.pos_name_aliases("Monge, David")
        self.assertIn("DAVID M", aliases)

    def test_short_date_label(self):
        self.assertEqual(schedule_app.short_date_label(date(2026, 7, 20)), "MON 7/20/26")

    def test_print_pages_balance_at_role_boundaries(self):
        grouped = [(object(), [object()] * count) for count in (9, 4, 13, 3, 4, 11, 4)]
        pages = schedule_app.balanced_print_pages(grouped)
        self.assertEqual([len(page) for page in pages], [3, 4])

    def test_print_view_has_full_short_dates_and_two_sheets(self):
        client = schedule_app.app.test_client()
        response = client.get("/print/2026-07-20")
        html = response.get_data(as_text=True)
        self.assertEqual(response.status_code, 200)
        self.assertIn("MON 7/20/26", html)
        self.assertIn("SUN 7/26/26", html)
        self.assertEqual(html.count('class="print-sheet"'), 2)


if __name__ == "__main__":
    unittest.main()
