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
