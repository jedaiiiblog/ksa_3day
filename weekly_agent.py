import os
import sys
from pathlib import Path
from datetime import datetime

sys.stdout.reconfigure(encoding="utf-8")

from dotenv import load_dotenv
from anthropic import Anthropic


SYSTEM_PROMPT = """당신은 여러 팀의 주간보고를 취합해 하나의 주간 경영 보고서 초안으로 정리하는 어시스턴트입니다.

아래 규칙을 반드시 지키세요.
- 제공된 원문에 없는 사실이나 수치를 만들어내지 마세요.
- 담당자나 기한이 명확하지 않으면 "확인 필요"라고 표시하세요.
- 같은 사안이 여러 팀 보고서에 등장하면 하나로 합치고, 관련 팀과 출처 파일명을 함께 적으세요.
- 모든 팀의 주요 내용이 누락되지 않도록 하세요.

결과는 아래 순서와 제목으로, 마크다운 형식으로 작성하세요.
1. 핵심 요약 (3줄)
2. 팀별 주요 실적
3. 여러 부서가 함께 확인할 이슈 (담당·기한 포함)
4. 대표가 결정해야 할 사항
5. 다음 주 주요 일정
"""


def load_api_key() -> str:
    load_dotenv()
    key = os.environ.get("ANTHROPIC_API_KEY")
    if not key or "여기에" in key:
        raise RuntimeError(".env 파일에 ANTHROPIC_API_KEY가 올바르게 설정되어 있지 않습니다.")
    return key


def load_team_reports(folder: str) -> list[tuple[str, str]]:
    folder_path = Path(folder)
    if not folder_path.is_dir():
        raise RuntimeError(f"'{folder}' 폴더를 찾을 수 없습니다.")

    files = sorted(
        f for f in folder_path.glob("*.md") if not f.name.startswith("주간경영보고_초안_생성")
    )
    if not files:
        raise RuntimeError(f"'{folder}' 폴더에서 .md 주간보고 파일을 찾지 못했습니다.")

    reports = []
    for f in files:
        text = f.read_text(encoding="utf-8").strip()
        if text:
            reports.append((f.name, text))

    if not reports:
        raise RuntimeError(f"'{folder}' 폴더의 파일 내용이 비어 있습니다.")

    return reports


def build_user_prompt(reports: list[tuple[str, str]]) -> str:
    parts = ["다음은 각 팀의 이번 주 보고서 원문입니다. 사안을 인용할 때 파일명을 출처로 사용하세요.\n"]
    for name, text in reports:
        parts.append(f"----- 파일: {name} -----\n{text}\n")
    return "\n".join(parts)


def resolve_output_path(folder: str) -> Path:
    today = datetime.now().strftime("%Y-%m-%d")
    output_dir = Path(folder)
    path = output_dir / f"주간경영보고_초안_생성_{today}.md"
    suffix = 1
    while path.exists():
        path = output_dir / f"주간경영보고_초안_생성_{today}_{suffix}.md"
        suffix += 1
    return path


def main() -> None:
    folder = sys.argv[1] if len(sys.argv) > 1 else "주간보고_연습_한빛정밀"

    try:
        api_key = load_api_key()
        reports = load_team_reports(folder)

        client = Anthropic(api_key=api_key)
        message = client.messages.create(
            model="claude-sonnet-5-5",
            max_tokens=8000,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": build_user_prompt(reports)}],
        )
        draft = next(
            (block.text for block in message.content if block.type == "text"), None
        )
        if not draft:
            raise RuntimeError(
                "모델 응답에서 텍스트를 받지 못했습니다 (stop_reason: "
                f"{message.stop_reason})."
            )
    except Exception as e:
        print(f"오류가 발생해 보고서를 생성하지 못했습니다: {e}")
        sys.exit(1)

    output_path = resolve_output_path(folder)
    output_path.write_text(draft, encoding="utf-8")

    print(f"읽은 팀 보고서 파일 수: {len(reports)}")
    print(f"결과 저장 위치: {output_path}")


if __name__ == "__main__":
    main()
