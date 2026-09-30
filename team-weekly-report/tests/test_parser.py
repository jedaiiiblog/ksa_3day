import unittest
from datetime import date

from team_report.parser import parse_report

SAMPLE = """\ufeff# 주간 보고
이름: 김민수
주차: 2026-09-25

## 프로젝트: A사 웹 개편
- 로그인 화면 완료
- 결제 연동 시작
## 프로젝트: B 시스템
- 배포 준비
"""


class ParseReportTest(unittest.TestCase):
    def test_basic(self):
        r = parse_report(SAMPLE)
        self.assertEqual(r.member, "김민수")
        self.assertEqual(r.week, date(2026, 9, 25))
        self.assertEqual([p.name for p in r.projects], ["A사 웹 개편", "B 시스템"])
        self.assertEqual(r.projects[0].body, "- 로그인 화면 완료\n- 결제 연동 시작")

    def test_crlf_and_fullwidth_colon(self):
        text = "이름： 김민수\r\n주차： 2026-09-25\r\n## 프로젝트： A\r\n- x\r\n"
        r = parse_report(text)
        self.assertEqual(r.member, "김민수")
        self.assertEqual(r.projects[0].name, "A")
        self.assertEqual(r.projects[0].body, "- x")

    def test_empty_body_project_still_counts(self):
        r = parse_report("이름: a\n주차: 2026-09-25\n## 프로젝트: A\n## 프로젝트: B\n- y\n")
        self.assertEqual([p.name for p in r.projects], ["A", "B"])
        self.assertEqual(r.projects[0].body, "")

    def test_duplicate_project_names_merge_ignoring_space_and_case(self):
        text = "이름: a\n주차: 2026-09-25\n## 프로젝트: Alpha  Web\n- 1\n## 프로젝트: alpha web\n- 2\n"
        r = parse_report(text)
        self.assertEqual(len(r.projects), 1)
        self.assertEqual(r.projects[0].name, "Alpha Web")
        self.assertEqual(r.projects[0].body, "- 1\n- 2")

    def test_missing_name_raises(self):
        with self.assertRaisesRegex(ValueError, "이름"):
            parse_report("주차: 2026-09-25\n## 프로젝트: A\n")

    def test_missing_week_raises(self):
        with self.assertRaisesRegex(ValueError, "주차"):
            parse_report("이름: a\n## 프로젝트: A\n")

    def test_bad_date_raises(self):
        with self.assertRaisesRegex(ValueError, "날짜"):
            parse_report("이름: a\n주차: 2026-13-45\n## 프로젝트: A\n")

    def test_no_projects_raises(self):
        with self.assertRaisesRegex(ValueError, "프로젝트"):
            parse_report("이름: a\n주차: 2026-09-25\n")


if __name__ == "__main__":
    unittest.main()
