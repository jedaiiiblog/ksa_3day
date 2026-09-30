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
