import hmac
import os
from datetime import datetime
from pathlib import Path

import streamlit as st

from proposal_agent import SECTIONS, check_sections, generate_proposal

SAMPLE_DIR = Path(__file__).parent / "주간보고_연습_한빛정밀"
MAX_FILES = 10
MAX_BYTES = 200_000  # 파일 하나당 최대 크기

st.set_page_config(page_title="공통 문제 해결방안 제안서", page_icon="📝", layout="wide")


def get_secret(name: str) -> str | None:
    try:
        return st.secrets[name]
    except Exception:
        return os.environ.get(name)


def require_password() -> None:
    """APP_PASSWORD가 설정된 경우에만 비밀번호를 요구한다. (API 사용료를 아무나 쓰지 못하게 하는 장치)"""
    expected = get_secret("APP_PASSWORD")
    if not expected or st.session_state.get("authed"):
        return
    pw = st.text_input("접속 비밀번호", type="password")
    if pw:
        if hmac.compare_digest(pw.encode(), expected.encode()):
            st.session_state["authed"] = True
            st.rerun()
        st.error("비밀번호가 맞지 않습니다.")
    st.stop()


def load_sample_reports() -> list[tuple[str, str]]:
    files = sorted(
        f for f in SAMPLE_DIR.glob("*.md") if not f.name.startswith("주간경영보고_초안_생성")
    )
    return [(f.name, f.read_text(encoding="utf-8").strip()) for f in files]


def read_uploads(files) -> list[tuple[str, str]]:
    reports = []
    for f in files:
        if f.size > MAX_BYTES:
            raise ValueError(f"'{f.name}' 파일이 너무 큽니다 (최대 {MAX_BYTES // 1000}KB).")
        try:
            text = f.getvalue().decode("utf-8").strip()
        except UnicodeDecodeError:
            raise ValueError(f"'{f.name}' 파일을 UTF-8 텍스트로 읽을 수 없습니다.")
        if text:
            reports.append((f.name, text))
    return reports


st.title("📝 공통 문제 해결방안 제안서")
st.caption(
    "여러 부서의 주간보고에서 공통 문제를 찾아 제안서를 만듭니다. "
    "근거 보고서를 함께 적고, 새 해결 방안은 [제안]으로, 근거가 부족한 내용은 '확인 필요'로 표시합니다."
)

require_password()

api_key = get_secret("ANTHROPIC_API_KEY")
if not api_key:
    st.error("ANTHROPIC_API_KEY가 설정되지 않았습니다. 앱 설정의 Secrets에 등록해 주세요.")
    st.stop()

source = st.radio(
    "보고서 선택",
    ["연습 자료 사용 (한빛정밀 5개 팀)", "내 보고서 파일 올리기"],
    horizontal=True,
)

reports: list[tuple[str, str]] = []
if source.startswith("연습"):
    reports = load_sample_reports()
    st.info(f"연습용 가상 자료 {len(reports)}건을 사용합니다: " + ", ".join(n for n, _ in reports))
else:
    st.warning("올린 보고서 내용은 제안서 작성을 위해 Anthropic API로 전송됩니다. 기밀 자료는 올리지 마세요.")
    uploads = st.file_uploader(
        f"주간보고 파일 (.md, .txt, 2개 이상, 최대 {MAX_FILES}개)",
        type=["md", "txt"],
        accept_multiple_files=True,
    )
    if uploads:
        if len(uploads) > MAX_FILES:
            st.error(f"파일은 최대 {MAX_FILES}개까지 올릴 수 있습니다.")
            st.stop()
        try:
            reports = read_uploads(uploads)
        except ValueError as e:
            st.error(str(e))
            st.stop()

if st.button("제안서 만들기", type="primary", disabled=len(reports) < 2):
    with st.spinner("보고서를 분석해 제안서를 작성하는 중입니다. 1~3분 걸릴 수 있습니다."):
        try:
            st.session_state["proposal"] = generate_proposal(reports, api_key)
        except Exception as e:
            st.session_state.pop("proposal", None)
            st.error(f"제안서를 만들지 못했습니다: {e}")
if len(reports) < 2:
    st.caption("공통 문제를 찾으려면 보고서가 2개 이상 필요합니다.")

proposal = st.session_state.get("proposal")
if proposal:
    for problem in check_sections(proposal):
        st.warning(f"형식 확인: {problem}")
    st.success("제안서가 만들어졌습니다. 아래 내용을 확인한 뒤 저장하세요.")
    st.download_button(
        "제안서 내려받기 (.md)",
        data=proposal.encode("utf-8"),
        file_name=f"공통문제_해결방안_제안서_{datetime.now():%Y-%m-%d}.md",
        mime="text/markdown",
    )
    st.divider()
    st.markdown(proposal)
