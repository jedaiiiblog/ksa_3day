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
