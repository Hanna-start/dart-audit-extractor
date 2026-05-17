# -*- coding: utf-8 -*-
"""
PreToolUse hook — CLAUDE.md의 I/O 규약을 시스템 레벨에서 강제한다.

차단 대상:
  1) `_input/raw/` 디렉토리에 대한 쓰기·수정·삭제 (Write/Edit/NotebookEdit/Bash)
  2) Bash 명령 중 비가역 파괴 패턴 (rm -rf, Remove-Item -Recurse -Force, del /s /q 등)

Claude Code hook protocol:
- stdin: JSON {"tool_name": "...", "tool_input": {...}}
- exit 0: 통과
- exit 2: 거부 (stderr 메시지 사용자에게 표시됨)
- 그 외 exit code: 에러로 간주되지만 작업은 진행
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

# stderr 한글 깨짐 방지 (사용자에게 차단 사유를 한글로 보여주기 위함)
try:
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass


# 프로젝트 루트(이 파일의 부모의 부모) 기준 절대 경로
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
RAW_DIR = (PROJECT_ROOT / "_input" / "raw").resolve()


def _is_under_raw(path_str: str) -> bool:
    """주어진 경로가 _input/raw/ 하위인지 — 상대/절대 모두 처리."""
    if not path_str:
        return False
    try:
        p = Path(path_str)
        if not p.is_absolute():
            p = (PROJECT_ROOT / p).resolve()
        else:
            p = p.resolve()
    except Exception:
        # 경로 파싱 실패 → 차단하지 않음 (false positive 방지)
        return False
    try:
        p.relative_to(RAW_DIR)
        return True
    except ValueError:
        return False


# 비가역 파괴 명령 패턴 (Bash 도구 입력 검사)
DESTRUCTIVE_PATTERNS = [
    re.compile(r"\brm\s+(-[rRfF]+\s+|--recursive\s+|--force\s+)"),       # rm -rf
    re.compile(r"\bRemove-Item\b.*-Recurse.*-Force", re.IGNORECASE),     # PS Remove-Item -Recurse -Force
    re.compile(r"\bdel\s+/[sSqQfF]", re.IGNORECASE),                     # cmd del /s /q
    re.compile(r"\brmdir\s+/[sSqQ]", re.IGNORECASE),                     # cmd rmdir /s /q
    re.compile(r"\bgit\s+clean\s+-[a-zA-Z]*[fF]"),                       # git clean -fd
    re.compile(r":\(\)\s*\{.*\};:"),                                     # fork bomb
]


def check_bash(command: str) -> tuple[bool, str]:
    """Bash 명령 검사. (allowed, reason)"""
    if not command:
        return True, ""

    # 비가역 파괴 명령
    for pat in DESTRUCTIVE_PATTERNS:
        if pat.search(command):
            return False, (
                f"비가역 파괴 명령 차단: 패턴 '{pat.pattern}' 매칭.\n"
                f"명령: {command[:200]}\n"
                f"CLAUDE.md의 raw 보호 원칙에 따라 거부됩니다."
            )

    # _input/raw/ 를 수정하려는 시도 (rm, mv, cp -f, >, >>, sed -i, tee 등)
    # 보수적으로: 명령에 _input/raw/ 가 등장하고, 동시에 쓰기 류 동사가 있으면 차단
    if "_input/raw" in command.replace("\\", "/").lower() or "_input\\raw" in command.lower():
        write_verbs = re.compile(
            r"\b(rm|del|move|mv|cp\s+-[fF]|copy\s+/y|Remove-Item|Move-Item|Out-File|"
            r"Set-Content|Add-Content|tee|sed\s+-i|>>?\s*[\"\']?.*_input)",
            re.IGNORECASE,
        )
        if write_verbs.search(command):
            return False, (
                f"_input/raw/ 영역에 대한 쓰기·삭제 시도 차단.\n"
                f"명령: {command[:200]}\n"
                f"raw는 read-only입니다. _workspace/ 또는 output/ 을 사용하세요."
            )

    return True, ""


def check_write_like(tool_input: dict) -> tuple[bool, str]:
    """Write/Edit/NotebookEdit 도구의 file_path 검사."""
    path = tool_input.get("file_path") or tool_input.get("notebook_path") or ""
    if _is_under_raw(path):
        return False, (
            f"_input/raw/ 영역의 파일 수정 차단: {path}\n"
            f"raw는 read-only입니다. _workspace/ 또는 output/ 을 사용하세요."
        )
    return True, ""


def main():
    try:
        payload = json.load(sys.stdin)
    except json.JSONDecodeError:
        # stdin 파싱 실패 시 통과 (hook 자체로 에러 일으키지 않음)
        sys.exit(0)

    tool = payload.get("tool_name", "")
    tool_input = payload.get("tool_input", {}) or {}

    if tool == "Bash":
        ok, reason = check_bash(tool_input.get("command", ""))
    elif tool in ("Write", "Edit", "NotebookEdit"):
        ok, reason = check_write_like(tool_input)
    else:
        ok, reason = True, ""

    if not ok:
        print(reason, file=sys.stderr)
        sys.exit(2)

    sys.exit(0)


if __name__ == "__main__":
    main()
