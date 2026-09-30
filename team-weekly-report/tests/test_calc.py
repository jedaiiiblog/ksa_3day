import unittest
from datetime import date

from team_report.calc import compute, workdays
from team_report.calendar_data import Calendar

HOL = {date(2026, 9, 24), date(2026, 9, 25)}


class WorkdaysTest(unittest.TestCase):
    def test_weekdays_only(self):
        self.assertEqual(workdays(2026, 9, set()), 22)

    def test_holidays_removed(self):
        self.assertEqual(workdays(2026, 9, HOL), 20)

    def test_vacation_removed(self):
        self.assertEqual(workdays(2026, 9, HOL, {date(2026, 9, 1), date(2026, 9, 2)}), 18)

    def test_weekend_holiday_and_vacation_not_double_counted(self):
        # 2026-09-26 토, 2026-09-27 일, 다른 달 날짜는 무시
        extra = {date(2026, 9, 26), date(2026, 9, 27), date(2026, 8, 31)}
        self.assertEqual(workdays(2026, 9, HOL | extra, extra), 20)

    def test_holiday_and_vacation_same_day_counted_once(self):
        self.assertEqual(workdays(2026, 9, HOL, {date(2026, 9, 24)}), 20)


class ComputeTest(unittest.TestCase):
    def setUp(self):
        self.cal = Calendar(
            holidays=set(HOL),
            vacations={
                "D": {date(2026, 9, 1), date(2026, 9, 2)},
                "E": {date(2026, 9, 1), date(2026, 9, 2), date(2026, 9, 3)},
            },
        )
        self.counts = {"A": 4, "B": 2, "C": 1, "D": 3, "E": 5}
        self.result = compute(self.counts, 2026, 9, self.cal)

    def test_spec_example(self):
        a = self.result["A"]
        self.assertEqual((a.workdays, a.base_hours), (20, 160))
        self.assertAlmostEqual(a.per_project_hours, 40)
        self.assertAlmostEqual(a.personal_pct, 25)
        self.assertAlmostEqual(a.team_share_pct, 160 / 760 * 100)
        self.assertAlmostEqual(a.project_share_pct, 40 / 760 * 100)

    def test_vacation_lowers_share(self):
        self.assertEqual(self.result["D"].base_hours, 144)
        self.assertEqual(self.result["E"].base_hours, 136)
        self.assertLess(self.result["E"].team_share_pct, self.result["A"].team_share_pct)

    def test_team_shares_sum_to_100(self):
        self.assertAlmostEqual(sum(c.team_share_pct for c in self.result.values()), 100)

    def test_project_shares_sum_to_team_share(self):
        for m, c in self.result.items():
            self.assertAlmostEqual(c.project_share_pct * self.counts[m], c.team_share_pct)

    def test_order_preserved(self):
        self.assertEqual(list(self.result), ["A", "B", "C", "D", "E"])

    def test_zero_projects_and_zero_team_hours_do_not_divide_by_zero(self):
        cal = Calendar(holidays={date(2026, 9, d) for d in range(1, 31)})
        r = compute({"A": 0, "B": 2}, 2026, 9, cal)
        self.assertEqual(r["A"].team_share_pct, 0.0)
        self.assertEqual(r["B"].per_project_hours, 0.0)
        r2 = compute({"A": 0}, 2026, 9, Calendar())
        self.assertEqual(r2["A"].per_project_hours, 0.0)
        self.assertEqual(r2["A"].personal_pct, 0.0)


if __name__ == "__main__":
    unittest.main()
