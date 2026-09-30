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
