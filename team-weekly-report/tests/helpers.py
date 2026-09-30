import os
from datetime import datetime
from pathlib import Path


def write_report(folder, filename, member, week, projects, saved_at):
    """projects: [(이름, 본문)]; saved_at: datetime (파일 mtime으로 설정)."""
    lines = ["# 주간 보고", f"이름: {member}", f"주차: {week}", ""]
    for name, body in projects:
        lines += [f"## 프로젝트: {name}", body]
    path = Path(folder) / filename
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    ts = saved_at.timestamp()
    os.utime(path, (ts, ts))
    return path
