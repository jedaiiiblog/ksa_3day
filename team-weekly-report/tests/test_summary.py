import json
import tempfile
import unittest
from pathlib import Path

from team_report.collector import MemberMonth
from team_report.summary import Summaries, export_input, load

MEMBERS = [MemberMonth("김민수", {"A사 웹 개편": ["- 로그인", "- 결제"], "B": ["- 배포"]})]


class SummaryTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.path = Path(self.tmp.name) / "요약_입력.json"

    def test_export_shape(self):
        export_input(MEMBERS, self.path)
        data = json.loads(self.path.read_text(encoding="utf-8"))
        self.assertEqual(data["팀요약"], "")
        self.assertEqual(data["팀원"]["김민수"]["A사 웹 개편"], {"원문": "- 로그인\n- 결제", "요약": ""})

    def test_export_refuses_to_overwrite_without_force(self):
        self.path.write_text("{}", encoding="utf-8")
        with self.assertRaises(FileExistsError):
            export_input(MEMBERS, self.path)
        self.assertEqual(self.path.read_text(encoding="utf-8"), "{}")  # 그대로
        export_input(MEMBERS, self.path, force=True)
        self.assertIn("김민수", self.path.read_text(encoding="utf-8"))

    def test_load_filled_summary_with_lenient_keys(self):
        export_input(MEMBERS, self.path)
        data = json.loads(self.path.read_text(encoding="utf-8"))
        data["팀요약"] = "팀 요약 문장"
        data["팀원"]["김민수"]["A사 웹 개편"]["요약"] = "로그인·결제 진행"
        self.path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
        s = load(self.path)
        self.assertEqual(s.team, "팀 요약 문장")
        self.assertEqual(s.get("김민수", "a사  웹 개편", ["x"]), "로그인·결제 진행")

    def test_blank_summary_falls_back_to_truncated_original(self):
        export_input(MEMBERS, self.path)
        s = load(self.path)
        self.assertEqual(s.get("김민수", "B", ["- 배포"]), "(AI 요약 없음) - 배포")
        long = s.get("김민수", "B", ["가" * 500])
        self.assertLessEqual(len(long), len("(AI 요약 없음) ") + 201)
        self.assertTrue(long.endswith("…"))

    def test_load_missing_file_is_empty(self):
        s = load(Path(self.tmp.name) / "없음.json")
        self.assertEqual((s.team, s.projects), ("", {}))

    def test_load_bad_json_raises_value_error(self):
        self.path.write_text("{not json", encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "요약"):
            load(self.path)


if __name__ == "__main__":
    unittest.main()
