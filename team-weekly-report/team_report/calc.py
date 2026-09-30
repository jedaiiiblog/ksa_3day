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
