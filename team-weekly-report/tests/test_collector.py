import tempfile
import unittest
from datetime import datetime, time
from pathlib import Path

from team_report.collector import collect, group_by_member
from tests.helpers import write_report

DEADLINE = time(17, 0)


class CollectTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)
        self.addCleanup(self.tmp.cleanup)

    def test_filters_month_and_sorts_by_saved_time(self):
        write_report(self.dir, "주간보고_b_2026-09-25.md", "B", "2026-09-25", [("P", "x")], datetime(2026, 9, 25, 15, 0))
        write_report(self.dir, "주간보고_a_2026-09-25.md", "A", "2026-09-25", [("P", "x")], datetime(2026, 9, 25, 10, 0))
        write_report(self.dir, "주간보고_a_2026-08-28.md", "A", "2026-08-28", [("P", "x")], datetime(2026, 8, 28, 10, 0))
        reports, problems = collect(self.dir, 2026, 9, DEADLINE)
        self.assertEqual([r.member for r in reports], ["A", "B"])
        self.assertEqual(problems, [])

    def test_late_flag_only_after_deadline(self):
        write_report(self.dir, "주간보고_a_1.md", "A", "2026-09-25", [("P", "x")], datetime(2026, 9, 25, 17, 0))
        write_report(self.dir, "주간보고_b_1.md", "B", "2026-09-25", [("P", "x")], datetime(2026, 9, 25, 17, 1))
        write_report(self.dir, "주간보고_c_1.md", "C", "2026-09-25", [("P", "x")], datetime(2026, 9, 28, 9, 0))
        reports, _ = collect(self.dir, 2026, 9, DEADLINE)
        self.assertEqual({r.member: r.late for r in reports}, {"A": False, "B": True, "C": True})

    def test_bad_file_becomes_problem_not_exception(self):
        (self.dir / "주간보고_x.md").write_text("이름: x\n", encoding="utf-8")
        (self.dir / "주간보고_y.md").write_bytes(b"\xff\xfe\x00bad")
        reports, problems = collect(self.dir, 2026, 9, DEADLINE)
        self.assertEqual(reports, [])
        self.assertEqual(len(problems), 2)
        self.assertTrue(any("주간보고_x.md" in p for p in problems))
        self.assertTrue(any("주간보고_y.md" in p for p in problems))

    def test_non_friday_week_warns_but_is_kept(self):
        write_report(self.dir, "주간보고_a.md", "A", "2026-09-23", [("P", "x")], datetime(2026, 9, 23, 9, 0))
        reports, problems = collect(self.dir, 2026, 9, DEADLINE)
        self.assertEqual(len(reports), 1)
        self.assertTrue(any("금요일" in p for p in problems))

    def test_ignores_unrelated_files(self):
        (self.dir / "index.html").write_text("<html>", encoding="utf-8")
        (self.dir / "팀_주간보고_취합_기획서.md").write_text("# 기획서", encoding="utf-8")
        reports, problems = collect(self.dir, 2026, 9, DEADLINE)
        self.assertEqual((reports, problems), ([], []))

    def test_txt_files_are_read_too(self):
        write_report(self.dir, "주간보고_a.txt", "A", "2026-09-25", [("P", "x")], datetime(2026, 9, 25, 9, 0))
        reports, _ = collect(self.dir, 2026, 9, DEADLINE)
        self.assertEqual(len(reports), 1)


class GroupTest(unittest.TestCase):
    def test_same_project_across_weeks_merges_ignoring_space_and_case(self):
        d = datetime(2026, 9, 25)
        with tempfile.TemporaryDirectory() as t:
            write_report(t, "주간보고_a_2.md", "A", "2026-09-25", [("alpha  web", "w2"), ("Beta", "b")], datetime(2026, 9, 25, 9))
            write_report(t, "주간보고_a_1.md", "A", "2026-09-18", [("Alpha Web", "w1")], datetime(2026, 9, 18, 9))
            reports, _ = collect(t, 2026, 9, time(17, 0))
        (m,) = group_by_member(reports)
        self.assertEqual(list(m.projects), ["Alpha Web", "Beta"])
        self.assertEqual(m.projects["Alpha Web"], ["w1", "w2"])  # 주차 순
        self.assertEqual(len(m.projects), 2)

    def test_member_order_follows_earliest_saved_time(self):
        with tempfile.TemporaryDirectory() as t:
            write_report(t, "주간보고_b_1.md", "B", "2026-09-18", [("P", "x")], datetime(2026, 9, 18, 9))
            write_report(t, "주간보고_a_1.md", "A", "2026-09-18", [("P", "x")], datetime(2026, 9, 18, 12))
            write_report(t, "주간보고_b_2.md", "B", "2026-09-25", [("P", "x")], datetime(2026, 9, 25, 16))
            write_report(t, "주간보고_a_2.md", "A", "2026-09-25", [("P", "x")], datetime(2026, 9, 25, 8))
            reports, _ = collect(t, 2026, 9, time(17, 0))
        self.assertEqual([m.member for m in group_by_member(reports)], ["B", "A"])


if __name__ == "__main__":
    unittest.main()
