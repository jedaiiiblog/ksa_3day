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
