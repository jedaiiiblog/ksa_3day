import re
import sys
from pathlib import Path
from datetime import datetime

# 어제 만든 weekly_agent.py는 수정하지 않고, 보고서 읽기/API 키 로딩 기능만 가져다 쓴다.
from weekly_agent import load_api_key, load_team_reports, build_user_prompt

from anthropic import Anthropic


SECTIONS = [
    "1. 현재 상황",
    "2. 확인된 문제",
    "3. 해결 방안 제안",
    "4. 기대 효과",
    "5. 필요한 것과 일정",
]

SYSTEM_PROMPT = """당신은 여러 부서의 주간보고를 읽고, 부서들에 공통으로 나타나는 문제를 찾아 해결 방안을 제안하는 제안서 작성 어시스턴트입니다.

아래 규칙을 반드시 지키세요.
- 문제를 적을 때마다 근거가 되는 보고서 파일명과 그 보고서의 해당 내용(짧게 인용 또는 요약)을 함께 적으세요. 근거 없는 문제는 쓰지 마세요.
- 2개 이상의 부서 보고서에서 확인되는 문제를 '공통 문제'로 다루고, 한 부서에서만 나온 문제는 그렇게 구분해서 표시하세요.
- 원문에 없는 새로운 해결 방안은 항목 앞에 반드시 "[제안]"이라고 표시하세요. 원문에 이미 적혀 있는 대응은 "[원문 내용]"으로 표시하고 파일명을 적으세요.
- 원문에 없는 효과 수치, 예산, 담당자, 확정 일정은 절대 만들지 마세요. 효과는 방향만 서술하고, 수치·예산·담당자·일정이 원문에 없으면 "확인 필요"라고 적으세요.
- 근거가 부족하거나 원문만으로 판단할 수 없는 내용은 "확인 필요"라고 표시하세요.
- 원문에 있는 날짜나 수치를 인용할 때는 원문 그대로 쓰세요.
- 원문에 없는 종합 평가, 분류, 해석("안정적", "간접 관련", "연쇄 영향", "~라는 뜻입니다" 등)은 사실처럼 쓰지 마세요. 꼭 필요하면 문장 앞에 "[해석]"을 붙이고, 원문의 어떤 내용에서 그렇게 읽었는지 적으세요. 원문에서 직접 확인되지 않는 부서 간 관련성은 "관련 여부 확인 필요"로 적으세요.
- 날짜는 원문이 정한 확정 수준(확정, 예정, 목표, 검토 중, 통보 예정)을 그대로 유지하세요. '예정'이나 '검토 중'인 날짜를 확정처럼 쓰지 마세요. 날짜의 의미(무슨 일정의 날짜인지)는 원문에 적힌 대로만 쓰세요.
- 날짜가 누구의 일정인지도 원문대로 적으세요. 다른 부서 보고서가 언급했더라도 그 날짜를 정한 부서는 원문에서 그 일정을 직접 적은 부서입니다.
- 서로 다른 단계의 날짜(예: 승인 목표, 초도 양산, 고객 납품)를 같은 종류의 날짜로 묶거나 '하나로 맞춘다'고 쓰지 마세요. 단계별 기준일을 구분해 관리한다고 쓰세요.
- 원문의 마감일이 무엇의 마감인지 바꾸지 마세요. 원문에 없는 마감 조건(예: 어떤 협상을 그 날짜까지 끝내야 한다)을 만들지 마세요.
- 제안이 원문의 계획을 바꾸는 내용이면(예: 발주 조건을 새로 정함) 원문 계획과 어떻게 다른지 함께 적고 "[제안]"으로 표시하세요. 담당자·통보 방식·절차를 정하는 문장도 원문에 없으면 모두 "[제안]"이며, 담당은 "확인 필요"입니다.
- 기대 효과는 검증된 사실이 아니라 제안이 실행될 경우의 기대라고 밝히고 "~할 수 있습니다" 수준으로만 쓰세요. 원문이 '필요하다'고만 한 일을 '누락됐다'고 쓰지 마세요.
- 원문에 있는 핵심 수량·금액·날짜(발주 규모, 필요 수량, 공지 수신일 등)는 관련 문제를 설명할 때 빠뜨리지 마세요. 그 수치가 예산 영향 분석의 기준인지는 "확인 필요"로 표시하세요.

결과는 아래 다섯 개 제목을 이 순서 그대로 마크다운 '## ' 제목으로 사용해 작성하세요. 다른 제목을 추가하거나 순서를 바꾸지 마세요.
## 1. 현재 상황
## 2. 확인된 문제
## 3. 해결 방안 제안
## 4. 기대 효과
## 5. 필요한 것과 일정
"""


def build_prompt(reports: list[tuple[str, str]]) -> str:
    return build_user_prompt(reports) + (
        "\n위 보고서들에서 부서 간 공통 문제를 찾아 제안서를 작성하세요. "
        "근거는 파일명을 사용하세요."
    )


def check_sections(text: str) -> list[str]:
    """필수 제목이 순서대로 모두 있는지 확인하고, 문제점 목록을 돌려준다."""
    problems = []
    last = -1
    for title in SECTIONS:
        m = re.search(rf"^##\s*{re.escape(title)}\s*$", text, re.MULTILINE)
        if not m:
            problems.append(f"'{title}' 제목이 없습니다.")
        elif m.start() < last:
            problems.append(f"'{title}' 제목의 순서가 바뀌었습니다.")
        else:
            last = m.start()
    return problems


def resolve_output_path(output_dir: Path) -> Path:
    today = datetime.now().strftime("%Y-%m-%d")
    path = output_dir / f"공통문제_해결방안_제안서_{today}.md"
    suffix = 1
    while path.exists():
        path = output_dir / f"공통문제_해결방안_제안서_{today}_{suffix}.md"
        suffix += 1
    return path


def main() -> None:
    folder = sys.argv[1] if len(sys.argv) > 1 else "주간보고_연습_한빛정밀"
    # 결과는 보고서 폴더와 분리해 저장한다. (같은 폴더에 두면 어제 프로그램이 제안서를 팀 보고서로 읽게 된다.)
    output_dir = Path(sys.argv[2] if len(sys.argv) > 2 else "제안서_결과")

    try:
        api_key = load_api_key()
        reports = load_team_reports(folder)

        client = Anthropic(api_key=api_key)
        message = client.messages.create(
            model="claude-sonnet-5-5",
            max_tokens=20000,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": build_prompt(reports)}],
        )
        if message.stop_reason == "max_tokens":
            raise RuntimeError("응답이 길이 제한에서 잘려 제안서가 완성되지 않았습니다.")
        proposal = next(
            (block.text for block in message.content if block.type == "text"), None
        )
        if not proposal:
            raise RuntimeError(
                f"모델 응답에서 텍스트를 받지 못했습니다 (stop_reason: {message.stop_reason})."
            )
    except Exception as e:
        print(f"오류가 발생해 제안서를 생성하지 못했습니다: {e}")
        sys.exit(1)

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = resolve_output_path(output_dir)
    output_path.write_text(proposal, encoding="utf-8")

    print(f"읽은 팀 보고서 파일 수: {len(reports)}")
    print(f"결과 저장 위치: {output_path.resolve()}")
    for p in check_sections(proposal):
        print(f"주의: {p}")


if __name__ == "__main__":
    main()
