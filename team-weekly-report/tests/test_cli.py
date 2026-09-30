import contextlib
import io
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from pptx import Presentation

from team_report.cli import main
from tests.helpers import write_report

CAL = "공휴일 2026-09-24 추석\n공휴일 2026-09-25 추석\n휴가 김민수 2026-09-01\n휴가 없는사람 2026-09-02\n"


def run(*argv):
    out, err = io.StringIO(), io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        try:
            code = main(list(argv))
        except SystemExit as e:
            code = e.code
    return code, out.getvalue() + err.getvalue()


class CliTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.dir = Path(self.tmp.name)
        write_report(self.dir, "주간보고_김민수_2026-09-25.md", "김민수", "2026-09-25",
                     [("A사 웹 개편", "- 로그인"), ("B", "- 배포")], datetime(2026, 9, 25, 10))
        write_report(self.dir, "주간보고_이서연_2026-09-25.md", "이서연", "2026-09-25",
                     [("A사 웹 개편", "- 디자인")], datetime(2026, 9, 28, 9))  # 마감 후
        (self.dir / "달력.txt").write_text(CAL, encoding="utf-8")

    def args(self, cmd, *extra):
        return (cmd, "--month", "2026-09", "--folder", str(self.dir), *extra)

    def test_check_reports_counts_late_and_warnings_and_writes_summary_input(self):
        code, out = run(*self.args("check"))
        self.assertEqual(code, 0, out)
        self.assertIn("김민수", out)
        self.assertIn("프로젝트 2개", out)
        self.assertIn("주간보고_이서연_2026-09-25.md", out)  # 지연 파일
        self.assertIn("없는사람", out)                        # 휴가 명단 이름 경고
        self.assertIn("5명", out)                             # 팀원 수 경고
        self.assertTrue((self.dir / "요약_입력.json").exists())

    def test_check_does_not_overwrite_filled_summary_without_force(self):
        self.assertEqual(run(*self.args("check"))[0], 0)
        p = self.dir / "요약_입력.json"
        p.write_text('{"팀요약": "AI가 쓴 것"}', encoding="utf-8")
        code, out = run(*self.args("check"))
        self.assertEqual(code, 2)
        self.assertIn("--force", out)
        self.assertIn("AI가 쓴 것", p.read_text(encoding="utf-8"))
        self.assertEqual(run(*self.args("check", "--force"))[0], 0)

    def test_build_creates_pptx_with_team_and_member_slides(self):
        run(*self.args("check"))
        code, out = run(*self.args("build"))
        self.assertEqual(code, 0, out)
        prs = Presentation(str(self.dir / "팀_월간보고_2026-09.pptx"))
        self.assertEqual(len(prs.slides), 3)

    def test_build_works_without_summary_file(self):
        code, _ = run(*self.args("build"))
        self.assertEqual(code, 0)

    def test_missing_calendar_aborts(self):
        (self.dir / "달력.txt").unlink()
        code, out = run(*self.args("check"))
        self.assertEqual(code, 2)
        self.assertIn("달력.txt", out)

    def test_bad_calendar_aborts_with_line_number(self):
        (self.dir / "달력.txt").write_text("휴가 김민수 2026-99-99\n", encoding="utf-8")
        code, out = run(*self.args("check"))
        self.assertEqual(code, 2)
        self.assertIn("달력 1행", out)

    def test_no_reports_for_month_aborts(self):
        code, out = run("check", "--month", "2026-10", "--folder", str(self.dir))
        self.assertEqual(code, 2)
        self.assertIn("2026-10", out)

    def test_invalid_month_and_deadline_rejected(self):
        self.assertEqual(run("check", "--month", "2026-13", "--folder", str(self.dir))[0], 2)
        self.assertEqual(run(*self.args("check", "--deadline", "25:99"))[0], 2)

    def test_deadline_option_changes_late_detection(self):
        code, out = run(*self.args("check", "--deadline", "23:59"))
        self.assertEqual(code, 0)
        # 이서연은 9/28 저장이라 여전히 지연, 김민수는 지연 아님
        self.assertIn("주간보고_이서연_2026-09-25.md", out)
        self.assertNotIn("지연: 주간보고_김민수", out)


if __name__ == "__main__":
    unittest.main()
