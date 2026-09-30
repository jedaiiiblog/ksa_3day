# 팀 주간 보고 취합 및 기여도 산출 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 실습 폴더의 팀원별 주간 보고 텍스트 파일을 월 단위로 취합하고, 근무일 기반 기여도를 코드로 계산해 상위부서 보고용 PPT(.pptx)를 만든다.

**Architecture:** 순수 Python 패키지 `team_report`. 파일 읽기·정렬·근무일·기여도·PPT 저장은 코드가 맡고, 요약 문장은 AI(Claude 세션)가 `요약_입력.json`의 빈 칸을 채우는 방식으로 분리한다. CLI는 두 단계다: `check`(취합 결과·경고 출력 + 요약 입력 파일 생성) → 사람 확인 + AI 요약 → `build`(PPT 생성). 기여도는 AI를 거치지 않는다.

**Tech Stack:** Python 3.14, python-pptx 1.0.2(설치됨), 표준 라이브러리 `unittest`(pytest 미설치이므로 추가 의존성 없음)

**Spec:** `팀_주간보고_취합_기획서.md`

> 이 폴더는 git 저장소가 아니다. 각 Task 끝의 "Commit" 단계는 `git init` 후에만 의미가 있으므로 생략하고 "테스트 통과 확인"으로 대신한다. 필요하면 실행 전에 `git init`을 먼저 한다.

## 기획서의 "확인 필요" 항목에 대한 임시 결정 (틀리면 이 표만 고치면 됨)

| 미정 항목 | 이 계획의 임시 결정 | 바꿀 때 영향 |
|---|---|---|
| 입력 파일 위치·이름 | `실습/주간보고_<이름>_<YYYY-MM-DD>.md` (또는 `.txt`), 파일 안에 `이름:`·`주차:` 줄 | Task 1, 4 |
| 마감 시각 | 주차 날짜(금요일) 17:00, `--deadline HH:MM`으로 변경 | Task 7 |
| 마감 후 저장 파일 | 제외하지 않고 "지연"으로 표시만 함 (사람이 판단) | Task 4, 6 |
| 공휴일·휴가 입력 | `달력.txt` (아래 형식), 하루 단위만 지원(반차 제외) | Task 2 |
| 월 내 여러 주 보고 | 프로젝트 이름의 합집합이 그 달 프로젝트 수 | Task 4 |
| "개인 기준 %" | 그 팀원 시간 중 해당 프로젝트 비중 = 100 ÷ 프로젝트 수 | Task 3 |
| 제출 주기 | 월말 1회 PPT (주간 PPT는 만들지 않음) | 전체 |

`달력.txt` 형식 (한 줄에 하나, `#` 뒤는 주석):

```
공휴일 2026-09-24 추석 연휴
휴가 김민수 2026-09-10
```

## Global Constraints

- 근무시간은 1인당 하루 8시간, 월~금 주 40시간.
- 개인 월 기준 시간 = (해당 월의 월~금 일수 − 공휴일 − 해당 팀원의 휴가일) × 8시간.
- 프로젝트별 시간 = 개인 월 기준 시간 ÷ 그 팀원의 프로젝트 개수.
- 팀 대비 비중(%) = 개인 월 기준 시간 ÷ 팀 월 기준 시간 합계. 프로젝트별 비중 = 프로젝트별 시간 ÷ 팀 합계.
- 기여도는 코드가 계산한다. AI는 요약·분류·문장 초안만 맡는다.
- 입력은 UTF-8 텍스트(.md/.txt). 저장 시각 순서는 파일 시스템 수정 시각(mtime)을 쓴다.
- 출력은 바탕화면 "실습" 폴더의 .pptx. 구성: 팀 전체 요약 1장 + 팀원별 1장(프로젝트 요약·기여도), 팀원 순서는 저장 시각 순.
- PPT에 반드시 포함: 팀원별 프로젝트 목록과 개수, 프로젝트별 요약, 월 기준 기여도(시간, 개인 기준 %, 팀 대비 비중 %).
- 하지 않는 것: 실투입 시간 반영, 달력 자동 연동, 기존 PPT 텍스트 추출, 입력 폼 화면, 자동 발송·알림, 순위·인사평가, 월 외 기간, 이미지·차트 슬라이드.
- 원본 파일은 읽기만 한다. 실습 폴더의 기존 파일(`index.html`, `공사스케줄_*.xlsx`, 기획서)은 건드리지 않는다.

## Review Focus

1. **`요약_입력.json` 덮어쓰기** — `check`를 다시 돌리면 AI가 채운 요약이 사라지기 쉽다. 기대 동작: 파일이 있으면 `--force` 없이는 중단하고 안내. (Task 5, 7)
2. **휴가 명단의 이름 오타 / 보고서 없는 팀원** — 이름이 다르면 휴가가 조용히 무시되어 시간이 과대 계산된다. 기대 동작: 경고를 출력. (Task 7)
3. **프로젝트 이름 표기 차이** — `A사 웹 개편`과 `A사  웹 개편`(공백)·대소문자 차이로 프로젝트 수가 부풀려지면 기여도가 틀어진다. 기대 동작: 공백·대소문자 무시하고 같은 프로젝트로 취합. (Task 1, 4)
4. **날짜가 금요일이 아니거나 저장 시각이 복사로 바뀐 파일** — 마감 판정이 틀린다. 기대 동작: 비금요일 주차 경고, 마감 후 저장은 "지연" 표시(제외 아님). (Task 4)
5. **프로젝트가 많거나 요약이 긴 팀원 슬라이드 / 달력 파일 없음** — 슬라이드가 깨지거나, 달력이 없을 때 공휴일 0으로 조용히 계산되면 안 된다. 기대 동작: 슬라이드는 생성되고(사람이 육안 검토), 달력이 없으면 오류로 중단. (Task 6, 7)

---

## File Structure

```
실습/
  team_report/
    __init__.py        # 빈 파일
    __main__.py        # python -m team_report 진입점
    models.py          # Project, Report, norm_name
    parser.py          # parse_report(text) -> Report
    calendar_data.py   # Calendar, parse_calendar(text)
    calc.py            # workdays, Contribution, compute
    collector.py       # collect(folder, year, month, deadline) / MemberMonth, group_by_member
    summary.py         # Summaries, export_input, load
    deck.py            # build_deck(...)
    cli.py             # check / build 명령
  tests/
    __init__.py
    helpers.py         # 테스트용 파일 생성 헬퍼
    test_parser.py  test_calendar.py  test_calc.py
    test_collector.py  test_summary.py  test_deck.py  test_cli.py
```

테스트 실행(전체): `python -m unittest discover -s tests -t . -v` (작업 디렉터리: `C:\Users\KSA\Desktop\실습`)

---

### Task 1: 모델과 보고서 파서

**Files:**
- Create: `team_report/__init__.py`, `tests/__init__.py` (둘 다 빈 파일)
- Create: `team_report/models.py`, `team_report/parser.py`
- Test: `tests/test_parser.py`

**Interfaces:**
- Consumes: 없음
- Produces:
  - `norm_name(name: str) -> str` — 공백 정리 + casefold (비교용 키)
  - `Project(name: str, body: str)` (frozen dataclass)
  - `Report(member: str, week: date, projects: list[Project], path: str = "", saved_at: datetime | None = None, late: bool = False)`
  - `parse_report(text: str) -> Report` — 형식 오류 시 `ValueError`(한국어 메시지)

입력 양식:

```
# 주간 보고
이름: 김민수
주차: 2026-09-25

## 프로젝트: A사 웹 개편
- 로그인 화면 완료
```

- [ ] **Step 1: 빈 `__init__.py` 두 개 생성 후 실패하는 테스트 작성**

`tests/test_parser.py`:

```python
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
```

- [ ] **Step 2: 실패 확인**

Run: `python -m unittest tests.test_parser -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'team_report.parser'`

- [ ] **Step 3: 구현**

`team_report/models.py`:

```python
from dataclasses import dataclass
from datetime import date, datetime


def norm_name(name: str) -> str:
    """프로젝트·팀원 이름 비교용 키: 공백 정리 + 대소문자 무시."""
    return " ".join(name.split()).casefold()


@dataclass(frozen=True)
class Project:
    name: str
    body: str


@dataclass
class Report:
    member: str
    week: date
    projects: list[Project]
    path: str = ""
    saved_at: datetime | None = None
    late: bool = False
```

`team_report/parser.py`:

```python
import re
from datetime import date

from .models import Project, Report, norm_name

_NAME = re.compile(r"^이름\s*[:：]\s*(.+)$")
_WEEK = re.compile(r"^주차\s*[:：]\s*(\S+)$")
_PROJECT = re.compile(r"^##\s*프로젝트\s*[:：]\s*(.+)$")


def parse_report(text: str) -> Report:
    member = week_raw = None
    display: dict[str, str] = {}
    bodies: dict[str, list[str]] = {}
    current = None
    for raw in text.lstrip("\ufeff").splitlines():
        line = raw.strip()
        m = _PROJECT.match(line)
        if m:
            name = " ".join(m.group(1).split())
            current = norm_name(name)
            display.setdefault(current, name)
            bodies.setdefault(current, [])
            continue
        if current is None:
            m = _NAME.match(line)
            if m:
                member = " ".join(m.group(1).split())
                continue
            m = _WEEK.match(line)
            if m:
                week_raw = m.group(1)
                continue
        elif line:
            bodies[current].append(line)
    if not member:
        raise ValueError("'이름:' 줄이 없습니다")
    if not week_raw:
        raise ValueError("'주차:' 줄이 없습니다")
    try:
        week = date.fromisoformat(week_raw)
    except ValueError:
        raise ValueError(f"주차 날짜 형식이 잘못되었습니다: {week_raw}") from None
    if not display:
        raise ValueError("'## 프로젝트:' 항목이 없습니다")
    projects = [Project(display[k], "\n".join(bodies[k])) for k in display]
    return Report(member, week, projects)
```

- [ ] **Step 4: 통과 확인**

Run: `python -m unittest tests.test_parser -v`
Expected: 8 tests OK

---

### Task 2: 달력(공휴일·휴가) 파서

**Files:**
- Create: `team_report/calendar_data.py`
- Test: `tests/test_calendar.py`

**Interfaces:**
- Consumes: 없음
- Produces:
  - `Calendar(holidays: set[date], vacations: dict[str, set[date]])` — vacations 키는 팀원 이름 그대로(공백 정리)
  - `parse_calendar(text: str) -> Calendar` — 잘못된 줄은 `ValueError("달력 N행: ...")`

- [ ] **Step 1: 실패하는 테스트**

`tests/test_calendar.py`:

```python
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
```

- [ ] **Step 2: 실패 확인**

Run: `python -m unittest tests.test_calendar -v`
Expected: FAIL — `ModuleNotFoundError`

- [ ] **Step 3: 구현**

`team_report/calendar_data.py`:

```python
from dataclasses import dataclass, field
from datetime import date


@dataclass
class Calendar:
    holidays: set[date] = field(default_factory=set)
    vacations: dict[str, set[date]] = field(default_factory=dict)


def parse_calendar(text: str) -> Calendar:
    cal = Calendar()
    for n, raw in enumerate(text.lstrip("\ufeff").splitlines(), 1):
        line = raw.split("#", 1)[0].strip()
        if not line:
            continue
        parts = line.split()
        try:
            if parts[0] == "공휴일" and len(parts) >= 2:
                cal.holidays.add(date.fromisoformat(parts[1]))
            elif parts[0] == "휴가" and len(parts) == 3:
                cal.vacations.setdefault(parts[1], set()).add(date.fromisoformat(parts[2]))
            else:
                raise ValueError("형식은 '공휴일 YYYY-MM-DD [이름]' 또는 '휴가 이름 YYYY-MM-DD'")
        except ValueError as e:
            raise ValueError(f"달력 {n}행: {line!r} ({e})") from None
    return cal
```

- [ ] **Step 4: 통과 확인**

Run: `python -m unittest tests.test_calendar -v`
Expected: 5 tests OK

---

### Task 3: 근무일·기여도 계산

**Files:**
- Create: `team_report/calc.py`
- Test: `tests/test_calc.py`

**Interfaces:**
- Consumes: `Calendar` (Task 2)
- Produces:
  - `HOURS_PER_DAY = 8`
  - `workdays(year: int, month: int, holidays: set[date], vacation_days: set[date] = frozenset()) -> int`
  - `Contribution(member, project_count, workdays, base_hours, per_project_hours, personal_pct, team_share_pct, project_share_pct)` — 앞 세 개는 int, 나머지는 float
  - `compute(project_counts: dict[str, int], year: int, month: int, cal: Calendar) -> dict[str, Contribution]` (입력 dict 순서 유지)

기준 검증값(2026-09): 월~금 22일, 공휴일 9/24·9/25 제외 시 20일 = 160h.

- [ ] **Step 1: 실패하는 테스트**

`tests/test_calc.py`:

```python
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
```

- [ ] **Step 2: 실패 확인**

Run: `python -m unittest tests.test_calc -v`
Expected: FAIL — `ModuleNotFoundError`

- [ ] **Step 3: 구현**

`team_report/calc.py`:

```python
import calendar as _cal
from dataclasses import dataclass
from datetime import date

from .calendar_data import Calendar

HOURS_PER_DAY = 8


def workdays(year: int, month: int, holidays: set[date], vacation_days: set[date] = frozenset()) -> int:
    last = _cal.monthrange(year, month)[1]
    count = 0
    for d in range(1, last + 1):
        day = date(year, month, d)
        if day.weekday() < 5 and day not in holidays and day not in vacation_days:
            count += 1
    return count


@dataclass(frozen=True)
class Contribution:
    member: str
    project_count: int
    workdays: int
    base_hours: int
    per_project_hours: float
    personal_pct: float       # 그 팀원 시간 중 프로젝트 1개의 비중 (100 / 프로젝트 수)
    team_share_pct: float     # 개인 월 기준 시간 / 팀 합계
    project_share_pct: float  # 프로젝트 1개 시간 / 팀 합계


def compute(project_counts: dict[str, int], year: int, month: int, cal: Calendar) -> dict[str, Contribution]:
    days = {
        m: workdays(year, month, cal.holidays, cal.vacations.get(m, set()))
        for m in project_counts
    }
    hours = {m: d * HOURS_PER_DAY for m, d in days.items()}
    total = sum(hours.values())
    out = {}
    for m, n in project_counts.items():
        per = hours[m] / n if n else 0.0
        out[m] = Contribution(
            member=m,
            project_count=n,
            workdays=days[m],
            base_hours=hours[m],
            per_project_hours=per,
            personal_pct=100 / n if n else 0.0,
            team_share_pct=hours[m] / total * 100 if total else 0.0,
            project_share_pct=per / total * 100 if total else 0.0,
        )
    return out
```

- [ ] **Step 4: 통과 확인**

Run: `python -m unittest tests.test_calc -v`
Expected: 11 tests OK

---

### Task 4: 파일 수집·정렬·지연 표시·팀원별 취합

**Files:**
- Create: `tests/helpers.py`, `team_report/collector.py`
- Test: `tests/test_collector.py`

**Interfaces:**
- Consumes: `parse_report`, `Report`, `norm_name` (Task 1)
- Produces:
  - `collect(folder: str | Path, year: int, month: int, deadline: time) -> tuple[list[Report], list[str]]` — 첫 값은 저장 시각(mtime) 오름차순 보고서 목록(각각 `path`, `saved_at`, `late` 채움), 둘째 값은 문제·경고 메시지 목록. 대상 파일: `주간보고_*.md|.txt`. 파싱 실패 파일은 예외 대신 메시지로.
  - `MemberMonth(member: str, projects: dict[str, list[str]], reports: list[Report])` — `projects`의 키는 표시 이름(처음 등장한 표기), 값은 주차 순 본문 목록
  - `group_by_member(reports: list[Report]) -> list[MemberMonth]` — 팀원 순서는 그 팀원의 가장 이른 저장 시각 순
- 마감 판정: `late = saved_at > datetime.combine(report.week, deadline)`

- [ ] **Step 1: 테스트 헬퍼와 실패하는 테스트**

`tests/helpers.py`:

```python
import os
from datetime import datetime
from pathlib import Path


def write_report(folder, filename, member, week, projects, saved_at):
    """projects: [(이름, 본문)]; saved_at: datetime (파일 mtime으로 설정)."""
    lines = ["# 주간 보고", f"이름: {member}", f"주차: {week}", ""]
    for name, body in projects:
        lines += [f"## 프로젝트: {name}", body]
    path = Path(folder) / filename
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    ts = saved_at.timestamp()
    os.utime(path, (ts, ts))
    return path
```

`tests/test_collector.py`:

```python
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
```

- [ ] **Step 2: 실패 확인**

Run: `python -m unittest tests.test_collector -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'team_report.collector'`

- [ ] **Step 3: 구현**

`team_report/collector.py`:

```python
from dataclasses import dataclass, field
from datetime import datetime, time
from pathlib import Path

from .models import Report, norm_name
from .parser import parse_report

_SUFFIXES = {".md", ".txt"}


def collect(folder, year: int, month: int, deadline: time) -> tuple[list[Report], list[str]]:
    reports: list[Report] = []
    problems: list[str] = []
    for p in sorted(Path(folder).iterdir()):
        if not (p.is_file() and p.name.startswith("주간보고_") and p.suffix.lower() in _SUFFIXES):
            continue
        try:
            r = parse_report(p.read_text(encoding="utf-8"))
        except (ValueError, UnicodeDecodeError) as e:
            problems.append(f"{p.name}: {e}")
            continue
        if (r.week.year, r.week.month) != (year, month):
            continue
        if r.week.weekday() != 4:
            problems.append(f"{p.name}: 주차 날짜 {r.week}가 금요일이 아닙니다 (마감 판정이 틀릴 수 있음)")
        r.path = str(p)
        r.saved_at = datetime.fromtimestamp(p.stat().st_mtime)
        r.late = r.saved_at > datetime.combine(r.week, deadline)
        reports.append(r)
    reports.sort(key=lambda r: r.saved_at)
    return reports, problems


@dataclass
class MemberMonth:
    member: str
    projects: dict[str, list[str]] = field(default_factory=dict)  # 표시 이름 -> 주차 순 본문
    reports: list[Report] = field(default_factory=list)


def group_by_member(reports: list[Report]) -> list[MemberMonth]:
    members: dict[str, MemberMonth] = {}
    for r in reports:  # 저장 시각 순 -> 팀원 순서 = 가장 이른 저장 순
        members.setdefault(r.member, MemberMonth(r.member)).reports.append(r)
    for mm in members.values():
        keys: dict[str, str] = {}  # norm -> 표시 이름
        for r in sorted(mm.reports, key=lambda r: (r.week, r.saved_at)):
            for proj in r.projects:
                name = keys.setdefault(norm_name(proj.name), proj.name)
                mm.projects.setdefault(name, []).append(proj.body)
    return list(members.values())
```

- [ ] **Step 4: 통과 확인**

Run: `python -m unittest tests.test_collector -v`
Expected: 8 tests OK

---

### Task 5: AI 요약 입출력 파일

**Files:**
- Create: `team_report/summary.py`
- Test: `tests/test_summary.py`

**Interfaces:**
- Consumes: `MemberMonth` (Task 4), `norm_name` (Task 1)
- Produces:
  - `Summaries(team: str = "", projects: dict[tuple[str, str], str] = {})` — 키는 `(norm_name(member), norm_name(project))`
  - `Summaries.get(member: str, project: str, bodies: list[str]) -> str` — 요약이 있으면 그것, 없으면 `"(AI 요약 없음) " + 원문 앞 200자`
  - `export_input(members: list[MemberMonth], path, force: bool = False) -> None` — 파일이 있고 `force=False`면 `FileExistsError`
  - `load(path) -> Summaries` — 파일이 없으면 빈 `Summaries`, JSON 오류는 `ValueError`

JSON 형식(AI가 각 `"요약"` 값과 `"팀요약"`을 채운다):

```json
{
  "안내": "각 '요약'과 '팀요약'을 상위부서 보고용 문장으로 채우세요. 숫자는 쓰지 마세요.",
  "팀요약": "",
  "팀원": {"김민수": {"A사 웹 개편": {"원문": "- 로그인 화면 완료", "요약": ""}}}
}
```

- [ ] **Step 1: 실패하는 테스트**

`tests/test_summary.py`:

```python
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
```

- [ ] **Step 2: 실패 확인**

Run: `python -m unittest tests.test_summary -v`
Expected: FAIL — `ModuleNotFoundError`

- [ ] **Step 3: 구현**

`team_report/summary.py`:

```python
import json
from dataclasses import dataclass, field
from pathlib import Path

from .collector import MemberMonth
from .models import norm_name

_FALLBACK_PREFIX = "(AI 요약 없음) "
_FALLBACK_LEN = 200


@dataclass
class Summaries:
    team: str = ""
    projects: dict[tuple[str, str], str] = field(default_factory=dict)

    def get(self, member: str, project: str, bodies: list[str]) -> str:
        found = self.projects.get((norm_name(member), norm_name(project)))
        if found:
            return found
        text = "\n".join(bodies)
        if len(text) > _FALLBACK_LEN:
            text = text[:_FALLBACK_LEN] + "…"
        return _FALLBACK_PREFIX + text


def export_input(members: list[MemberMonth], path, force: bool = False) -> None:
    path = Path(path)
    if path.exists() and not force:
        raise FileExistsError(str(path))
    data = {
        "안내": "각 '요약'과 '팀요약'을 상위부서 보고용 문장으로 채우세요. 숫자(시간·%)는 쓰지 마세요.",
        "팀요약": "",
        "팀원": {
            mm.member: {
                name: {"원문": "\n".join(bodies), "요약": ""}
                for name, bodies in mm.projects.items()
            }
            for mm in members
        },
    }
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def load(path) -> Summaries:
    path = Path(path)
    if not path.exists():
        return Summaries()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        out = Summaries(team=str(data.get("팀요약", "")).strip())
        for member, projects in data.get("팀원", {}).items():
            for name, entry in projects.items():
                text = str(entry.get("요약", "")).strip()
                if text:
                    out.projects[(norm_name(member), norm_name(name))] = text
        return out
    except (json.JSONDecodeError, AttributeError, UnicodeDecodeError) as e:
        raise ValueError(f"요약 파일을 읽을 수 없습니다: {path.name} ({e})") from None
```

- [ ] **Step 4: 통과 확인**

Run: `python -m unittest tests.test_summary -v`
Expected: 6 tests OK

---

### Task 6: PPT 생성

**Files:**
- Create: `team_report/deck.py`
- Test: `tests/test_deck.py`

**Interfaces:**
- Consumes: `MemberMonth` (Task 4), `Contribution` (Task 3), `Summaries` (Task 5)
- Produces: `build_deck(out_path, year: int, month: int, members: list[MemberMonth], contribs: dict[str, Contribution], summaries: Summaries) -> None`
  - 슬라이드 1: 제목 `"{year}년 {month}월 팀 주간 보고 취합"`, 표(팀원/프로젝트 수/월 기준 시간/팀 대비 비중), 하단에 `summaries.team`(없으면 "(AI 요약 없음)")
  - 슬라이드 2~: 팀원별. 제목 `"{member} — {year}년 {month}월"`, 표(프로젝트/요약/시간/개인 기준 %/팀 대비 %), 하단 메모: 근무일·월 기준 시간, 지연 저장 파일이 있으면 그 파일명
  - 시간은 소수 1자리(`53.3h`), 비율은 소수 1자리(`21.1%`)

- [ ] **Step 1: 실패하는 테스트**

`tests/test_deck.py`:

```python
import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path

from pptx import Presentation

from team_report.calc import compute
from team_report.calendar_data import Calendar
from team_report.collector import MemberMonth
from team_report.deck import build_deck
from team_report.models import Project, Report
from team_report.summary import Summaries

HOL = {date(2026, 9, 24), date(2026, 9, 25)}


def slide_text(slide):
    parts = []
    for sh in slide.shapes:
        if sh.has_text_frame:
            parts.append(sh.text_frame.text)
        if sh.has_table:
            for row in sh.table.rows:
                parts.append(" | ".join(c.text for c in row.cells))
    return "\n".join(parts)


class DeckTest(unittest.TestCase):
    def build(self, members, summaries=None):
        counts = {m.member: len(m.projects) for m in members}
        contribs = compute(counts, 2026, 9, Calendar(holidays=set(HOL)))
        d = tempfile.TemporaryDirectory()
        self.addCleanup(d.cleanup)
        out = Path(d.name) / "out.pptx"
        build_deck(out, 2026, 9, members, contribs, summaries or Summaries(team="팀 요약"))
        return Presentation(str(out))

    def test_slide_count_order_and_numbers(self):
        a = MemberMonth("A", {f"P{i}": ["x"] for i in range(4)})
        b = MemberMonth("B", {"Q": ["y"]})
        prs = self.build([a, b])
        self.assertEqual(len(prs.slides), 3)
        team = slide_text(prs.slides[0])
        self.assertIn("2026년 9월", team)
        self.assertIn("팀 요약", team)
        self.assertIn("50.0%", team)  # A, B 모두 160h -> 각 50%
        sa = slide_text(prs.slides[1])
        self.assertIn("A — 2026년 9월", sa)
        self.assertIn("40.0h", sa)     # 160 / 4
        self.assertIn("25.0%", sa)     # 개인 기준
        self.assertIn("12.5%", sa)     # 40 / 320
        self.assertIn("근무일 20일", sa)
        self.assertIn("160.0h", sa)
        self.assertIn("B — 2026년 9월", slide_text(prs.slides[2]))

    def test_summary_used_and_fallback_marked(self):
        a = MemberMonth("A", {"P1": ["원문1"], "P2": ["원문2"]})
        s = Summaries(team="t", projects={("a", "p1"): "요약된 문장"})
        text = slide_text(self.build([a], s).slides[1])
        self.assertIn("요약된 문장", text)
        self.assertIn("(AI 요약 없음) 원문2", text)

    def test_late_files_listed_on_member_slide(self):
        r = Report("A", date(2026, 9, 25), [Project("P", "x")], path="C:/x/주간보고_a_1.md",
                   saved_at=datetime(2026, 9, 28, 9), late=True)
        a = MemberMonth("A", {"P": ["x"]}, [r])
        self.assertIn("주간보고_a_1.md", slide_text(self.build([a]).slides[1]))

    def test_member_with_many_projects_and_long_text_still_builds(self):
        a = MemberMonth("A", {f"프로젝트{i}": ["가" * 400] for i in range(12)})
        prs = self.build([a])
        self.assertEqual(len(prs.slides), 2)

    def test_no_members_builds_team_slide_only(self):
        self.assertEqual(len(self.build([]).slides), 1)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: 실패 확인**

Run: `python -m unittest tests.test_deck -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'team_report.deck'`

- [ ] **Step 3: 구현**

`team_report/deck.py`:

```python
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt

from .calc import Contribution
from .collector import MemberMonth
from .summary import Summaries

_FONT = "맑은 고딕"
_TITLE_ONLY = 5


def _style(text_frame, size, bold=False):
    for p in text_frame.paragraphs:
        for r in p.runs:
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.name = _FONT


def _fill(cell, text, size, bold=False):
    cell.text_frame.word_wrap = True
    cell.text_frame.text = text
    _style(cell.text_frame, size, bold)


def _table(slide, rows, widths, top, size):
    shape = slide.shapes.add_table(
        len(rows), len(rows[0]), Inches(0.5), Inches(top), Inches(sum(widths)), Inches(0.4 * len(rows))
    )
    for i, w in enumerate(widths):
        shape.table.columns[i].width = Inches(w)
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            _fill(shape.table.cell(r, c), str(value), size, bold=(r == 0))
    return shape


def _note(slide, text, top, size=12):
    box = slide.shapes.add_textbox(Inches(0.5), Inches(top), Inches(12.3), Inches(1.0))
    box.text_frame.word_wrap = True
    box.text_frame.text = text
    _style(box.text_frame, size)


def _new_slide(prs, title):
    slide = prs.slides.add_slide(prs.slide_layouts[_TITLE_ONLY])
    slide.shapes.title.text = title
    _style(slide.shapes.title.text_frame, 28, bold=True)
    return slide


def _pct(v):
    return f"{v:.1f}%"


def _hours(v):
    return f"{v:.1f}h"


def build_deck(out_path, year: int, month: int, members: list[MemberMonth],
               contribs: dict[str, Contribution], summaries: Summaries) -> None:
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)

    team = _new_slide(prs, f"{year}년 {month}월 팀 주간 보고 취합")
    rows = [["팀원", "프로젝트 수", "월 기준 시간", "팀 대비 비중"]]
    for mm in members:
        c = contribs[mm.member]
        rows.append([mm.member, c.project_count, _hours(c.base_hours), _pct(c.team_share_pct)])
    _table(team, rows, [3.5, 2.6, 3.1, 3.1], 1.5, 14)
    _note(team, summaries.team or "(AI 요약 없음)", 1.5 + 0.4 * len(rows) + 0.4, 14)

    for mm in members:
        c = contribs[mm.member]
        slide = _new_slide(prs, f"{mm.member} — {year}년 {month}월")
        rows = [["프로젝트", "요약", "시간", "개인 기준 %", "팀 대비 %"]]
        for name, bodies in mm.projects.items():
            rows.append([
                name,
                summaries.get(mm.member, name, bodies),
                _hours(c.per_project_hours),
                _pct(c.personal_pct),
                _pct(c.project_share_pct),
            ])
        size = 12 if len(rows) <= 6 else 9
        _table(slide, rows, [2.4, 5.6, 1.1, 1.6, 1.6], 1.4, size)
        memo = f"근무일 {c.workdays}일 · 월 기준 시간 {_hours(c.base_hours)} · 팀 대비 {_pct(c.team_share_pct)}"
        late = [Path(r.path).name for r in mm.reports if r.late]
        if late:
            memo += "\n마감 후 저장된 파일: " + ", ".join(late)
        _note(slide, memo, 6.5, 11)

    prs.save(str(out_path))
```

- [ ] **Step 4: 통과 확인**

Run: `python -m unittest tests.test_deck -v`
Expected: 5 tests OK

- [ ] **Step 5: 육안 확인(사람)**

12개 프로젝트 테스트 결과물을 PowerPoint에서 열어 글자가 넘치는지 본다. 넘치면 `size`·표 높이 기준을 조정한다(자동 검증 불가 — 기획서의 "사람이 확인" 항목에 해당).

---

### Task 7: CLI(check / build)와 종단 테스트

**Files:**
- Create: `team_report/cli.py`, `team_report/__main__.py`
- Test: `tests/test_cli.py`

**Interfaces:**
- Consumes: `collect`, `group_by_member` (Task 4), `parse_calendar` (Task 2), `compute` (Task 3), `export_input`, `load` (Task 5), `build_deck` (Task 6)
- Produces: `main(argv: list[str] | None = None) -> int` (0 성공, 2 오류). 명령:
  - `python -m team_report check --month YYYY-MM [--folder .] [--deadline 17:00] [--force]`
    → 취합 결과 출력(팀원·프로젝트 목록·개수·근무일·시간, 지연 파일, 문제·경고), `요약_입력.json` 생성
  - `python -m team_report build --month YYYY-MM [--folder .] [--deadline 17:00]`
    → `팀_월간보고_YYYY-MM.pptx` 저장
- 고정 파일명(폴더 안): `달력.txt`, `요약_입력.json`
- 중단(코드 2) 조건: 달력 파일 없음·형식 오류, 해당 월 보고서 0건, `check`에서 요약 파일이 있는데 `--force` 없음
- 경고(계속 진행): 팀원 수가 5명이 아님, 휴가 명단에 보고서 없는 이름, 파싱 실패·비금요일 파일

- [ ] **Step 1: 실패하는 테스트**

`tests/test_cli.py`:

```python
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
```

- [ ] **Step 2: 실패 확인**

Run: `python -m unittest tests.test_cli -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'team_report.cli'`

- [ ] **Step 3: 구현**

`team_report/cli.py`:

```python
import argparse
import sys
from datetime import date, time
from pathlib import Path

from .calc import compute
from .calendar_data import parse_calendar
from .collector import collect, group_by_member
from .deck import build_deck
from .summary import export_input, load

CALENDAR_FILE = "달력.txt"
SUMMARY_FILE = "요약_입력.json"
TEAM_SIZE = 5


def _month(s: str):
    try:
        y, m = (int(x) for x in s.split("-"))
        date(y, m, 1)
    except ValueError:
        raise argparse.ArgumentTypeError("YYYY-MM 형식으로 입력하세요 (예: 2026-09)") from None
    return y, m


def _deadline(s: str) -> time:
    try:
        return time.fromisoformat(s)
    except ValueError:
        raise argparse.ArgumentTypeError("HH:MM 형식으로 입력하세요 (예: 17:00)") from None


def _parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="team_report", description="팀 주간 보고 월간 취합·기여도 PPT")
    sub = p.add_subparsers(dest="cmd", required=True)
    for name in ("check", "build"):
        s = sub.add_parser(name)
        s.add_argument("--month", type=_month, required=True)
        s.add_argument("--folder", default=".")
        s.add_argument("--deadline", type=_deadline, default=time(17, 0))
        if name == "check":
            s.add_argument("--force", action="store_true", help="기존 요약_입력.json을 덮어쓴다")
    return p


def _prepare(args):
    """공통 준비. 오류면 (None, 메시지)."""
    folder = Path(args.folder)
    year, month = args.month
    cal_path = folder / CALENDAR_FILE
    if not cal_path.exists():
        return None, f"{CALENDAR_FILE}이 없습니다: {cal_path} (공휴일·휴가를 적은 파일이 필요합니다)"
    try:
        cal = parse_calendar(cal_path.read_text(encoding="utf-8"))
    except (ValueError, UnicodeDecodeError) as e:
        return None, f"{CALENDAR_FILE}: {e}"
    reports, problems = collect(folder, year, month, args.deadline)
    if not reports:
        return None, f"{year}-{month:02d}에 해당하는 주간 보고 파일이 없습니다 ({folder})" + "".join(f"\n  - {p}" for p in problems)
    members = group_by_member(reports)
    contribs = compute({m.member: len(m.projects) for m in members}, year, month, cal)
    warnings = list(problems)
    if len(members) != TEAM_SIZE:
        warnings.append(f"팀원이 {len(members)}명입니다 (기획서 기준 {TEAM_SIZE}명). 팀 합계와 비중이 달라집니다")
    known = {m.member for m in members}
    for name in cal.vacations:
        if name not in known:
            warnings.append(f"휴가 명단의 '{name}'은(는) 보고서에 없는 이름입니다 (오타 확인)")
    return (year, month, folder, members, contribs, warnings), None


def _print_check(year, month, members, contribs, warnings):
    print(f"[{year}-{month:02d}] 취합 결과 (저장 시각 순)")
    for mm in members:
        c = contribs[mm.member]
        print(f"- {mm.member}: 프로젝트 {c.project_count}개 ({', '.join(mm.projects)}) / "
              f"근무일 {c.workdays}일 {c.base_hours}h / 팀 대비 {c.team_share_pct:.1f}%")
        for r in mm.reports:
            if r.late:
                print(f"    지연: {Path(r.path).name} (저장 {r.saved_at:%Y-%m-%d %H:%M})")
    for w in warnings:
        print(f"경고: {w}")
    print("확인할 것: ① 프로젝트 개수 ② 공휴일·휴가 옮겨 적기 ③ 기여도 계산 결과")


def main(argv=None) -> int:
    try:
        args = _parser().parse_args(argv)
    except SystemExit as e:
        return 2 if e.code else 0
    prepared, error = _prepare(args)
    if error:
        print(f"오류: {error}", file=sys.stderr)
        return 2
    year, month, folder, members, contribs, warnings = prepared
    if args.cmd == "check":
        try:
            export_input(members, folder / SUMMARY_FILE, force=args.force)
        except FileExistsError:
            print(f"오류: {SUMMARY_FILE}이 이미 있습니다. AI가 채운 요약이 사라질 수 있으니 "
                  "덮어쓰려면 --force를 붙이세요.", file=sys.stderr)
            return 2
        _print_check(year, month, members, contribs, warnings)
        print(f"요약 입력 파일 생성: {folder / SUMMARY_FILE}")
        return 0
    try:
        summaries = load(folder / SUMMARY_FILE)
    except ValueError as e:
        print(f"오류: {e}", file=sys.stderr)
        return 2
    out = folder / f"팀_월간보고_{year}-{month:02d}.pptx"
    try:
        build_deck(out, year, month, members, contribs, summaries)
    except PermissionError:
        print(f"오류: {out.name}을 쓸 수 없습니다. PowerPoint에서 열려 있으면 닫고 다시 실행하세요.", file=sys.stderr)
        return 2
    for w in warnings:
        print(f"경고: {w}")
    print(f"PPT 저장: {out}")
    return 0
```

`team_report/__main__.py`:

```python
import sys

from .cli import main

sys.exit(main())
```

- [ ] **Step 4: 통과 확인**

Run: `python -m unittest tests.test_cli -v`
Expected: 9 tests OK

- [ ] **Step 5: 전체 테스트**

Run: `python -m unittest discover -s tests -t . -v`
Expected: 모든 테스트 OK (parser 8, calendar 5, calc 11, collector 8, summary 6, deck 5, cli 9)

- [ ] **Step 6: 실제 폴더 리허설 (사람 확인 포함)**

1. 실습 폴더에 샘플 `주간보고_*.md` 5개와 `달력.txt`를 만든다(테스트 데이터, 실제 팀원 데이터 아님).
2. `python -m team_report check --month 2026-09` → 콘솔의 프로젝트 개수·지연·경고 확인
3. 생성된 `요약_입력.json`의 `"요약"`·`"팀요약"`을 Claude가 채운다(숫자 제외).
4. `python -m team_report build --month 2026-09` → `팀_월간보고_2026-09.pptx`를 팀 취합 담당자가 검토한 뒤 제출.

---

## Self-Review

- **Spec coverage:** 파일 찾기·팀원별 선택·저장 시각 정렬·마감 후 표시(Task 4), 파일 읽기(Task 1), 프로젝트 개수(Task 4 `group_by_member`), 근무일 계산(공휴일·휴가 제외, Task 3), 기여도(Task 3), PPT 저장(Task 6), AI 요약 분리(Task 5), 사람 확인 4항목(Task 7 `check` 출력 + Task 6 Step 5 + 리허설), "이번에 안 할 것"은 구현하지 않음. 기획서 §7 계산 예시(160h → 21.1%)는 Task 3 `test_spec_example`이 검증.
- **Placeholder scan:** 없음.
- **Type consistency:** `Contribution` 필드명(`workdays`, `base_hours`, `per_project_hours`, `personal_pct`, `team_share_pct`, `project_share_pct`)은 Task 3 정의·Task 6/7 사용에서 일치. `MemberMonth.projects`(표시 이름 → 본문 목록)는 Task 4·5·6에서 동일. `collect()` 반환 `(reports, problems)`는 Task 4·7에서 동일.
- **Known gaps:** 반차(하루 미만 휴가), 휴가 기간 범위 입력(`~`), 다른 이름 표기(예: 성만 쓰기)는 미지원. 기획서 "확인 필요" 항목이 확정되면 상단 결정 표를 고치고 해당 Task를 수정한다.
