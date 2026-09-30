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
