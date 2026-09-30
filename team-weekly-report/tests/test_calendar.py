import unittest
from datetime import date

from team_report.calendar_data import parse_calendar

TEXT = """\ufeff# 9월
공휴일 2026-09-24 추석 연휴
공휴일 2026-09-25
휴가 김민수 2026-09-10   # 개인 휴가
휴가 김민수 2026-09-11

"""


class ParseCalendarTest(unittest.TestCase):
    def test_parses_holidays_and_vacations(self):
        c = parse_calendar(TEXT)
        self.assertEqual(c.holidays, {date(2026, 9, 24), date(2026, 9, 25)})
        self.assertEqual(c.vacations, {"김민수": {date(2026, 9, 10), date(2026, 9, 11)}})

    def test_empty_text_is_empty_calendar(self):
        c = parse_calendar("")
        self.assertEqual((c.holidays, c.vacations), (set(), {}))

    def test_bad_date_reports_line_number(self):
        with self.assertRaisesRegex(ValueError, "달력 2행"):
            parse_calendar("공휴일 2026-09-24\n휴가 김민수 2026-99-99\n")

    def test_unknown_keyword_raises(self):
        with self.assertRaisesRegex(ValueError, "달력 1행"):
            parse_calendar("연차 김민수 2026-09-10\n")

    def test_vacation_without_name_raises(self):
        with self.assertRaisesRegex(ValueError, "달력 1행"):
            parse_calendar("휴가 2026-09-10\n")


if __name__ == "__main__":
    unittest.main()
