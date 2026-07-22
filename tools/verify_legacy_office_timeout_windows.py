"""Auditable Windows-only check for legacy Office timeout cleanup.

Run against a disposable DOC/XLS/PPT copy.  It records only Office PIDs that
exist before and after this task and never kills a PID discovered from the
user's desktop session.
"""
from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

from src.legacy_office import extract_legacy_readonly


def office_pids() -> set[int]:
    output = subprocess.check_output(["tasklist", "/FO", "CSV", "/NH"], text=True, encoding="mbcs", errors="replace")
    names = {"WINWORD.EXE", "EXCEL.EXE", "POWERPNT.EXE"}; found: set[int] = set()
    for line in output.splitlines():
        fields = [part.strip('"') for part in line.split('","')]
        if len(fields) >= 2 and fields[0].upper() in names:
            try: found.add(int(fields[1]))
            except ValueError: pass
    return found


def main() -> int:
    parser = argparse.ArgumentParser(description="审计旧版 Office 超时后仅清理本任务创建的 PID")
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--timeout", type=int, default=1)
    args = parser.parse_args()
    if sys.platform != "win32": parser.error("此审计脚本仅适用于 Windows")
    before = office_pids(); print(f"启动前 Office PID：{sorted(before)}")
    try: extract_legacy_readonly(args.input, args.timeout)
    except Exception as exc: print(f"转换结果：{type(exc).__name__}: {exc}")
    after = office_pids(); print(f"结束后 Office PID：{sorted(after)}")
    leaked = after - before
    if leaked:
        print(f"检测到残留 Office PID（需要人工调查）：{sorted(leaked)}")
        return 2
    print("未检测到本次任务遗留的 Office PID。")
    return 0


if __name__ == "__main__": raise SystemExit(main())
