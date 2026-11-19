#!/usr/bin/env python3
"""PreToolUse 훅: 원본·보존 폴더를 건드리거나 위험한 명령을 쓰려 하면 막는다.

Claude Code는 도구 실행 직전 이 스크립트에 JSON(stdin)을 넘긴다.
종료 코드 2 = 차단(stderr 내용이 Claude에게 전달됨), 0 = 통과.
보호 폴더는 환경변수 RECORDS_PROTECTED (쉼표 구분)로 바꿀 수 있다.
"""
import json
import os
import re
import sys

PROTECTED = [p.strip().strip("/") for p in (os.environ.get("RECORDS_PROTECTED") or os.environ.get("CLAUDE_PLUGIN_OPTION_PROTECTED_FOLDERS") or "원본,보존,originals,preserved").split(",") if p.strip()]
# 비밀기록물 폴더: 읽기조차 막는다(AI 처리 대상이 아님)
SECRET = [p.strip().strip("/") for p in (os.environ.get("RECORDS_SECRET") or os.environ.get("CLAUDE_PLUGIN_OPTION_SECRET_FOLDERS") or "비밀,secret,classified").split(",") if p.strip()]

# 실행 자체를 막을 명령 (삭제·외부 전송·권한 변경)
DANGEROUS = [
    (r"(^|[\s;&|])rm\s", "파일 삭제(rm)"),
    (r"(^|[\s;&|])rmdir\s", "폴더 삭제(rmdir)"),
    (r"(^|[\s;&|])(shred|unlink)\s", "파일 삭제"),
    (r"(^|[\s;&|])(curl|wget|scp|rsync|ftp|sftp)\s", "외부 전송"),
    (r"(^|[\s;&|])chmod\s", "권한 변경"),
    (r"git\s+push", "원격 저장소 전송"),
]


def touches_protected(text: str, names=None) -> str | None:
    norm = text.replace("\\", "/")
    for p in (names or PROTECTED):
        if re.search(rf"(^|[\s'\"/=]){re.escape(p)}(/|$|[\s'\"])", norm):
            return p
    return None


def block(msg: str) -> None:
    print(f"[기록관리 하네스] 차단: {msg}", file=sys.stderr)
    sys.exit(2)


def main() -> None:
    try:
        data = json.load(sys.stdin)
    except json.JSONDecodeError:
        sys.exit(0)  # 형식을 모르면 판단하지 않는다(다른 규칙이 처리)
    tool = data.get("tool_name", "")
    inp = data.get("tool_input", {}) or {}

    # 비밀기록물: 어떤 도구로도 읽거나 쓰지 않는다
    probe = " ".join(str(inp.get(k, "")) for k in ("file_path", "notebook_path", "path", "pattern", "command"))
    hit = touches_protected(probe, SECRET)
    if hit:
        block(f"'{hit}' 폴더는 비밀기록물 보관 위치로 보고 AI 처리에서 제외합니다. 비밀기록물은 담당자가 직접 다루세요.")

    if tool in ("Write", "Edit", "MultiEdit", "NotebookEdit"):
        path = inp.get("file_path") or inp.get("notebook_path") or ""
        hit = touches_protected(path)
        if hit:
            block(f"'{hit}' 폴더는 읽기 전용입니다. 결과는 새 파일로 다른 폴더에 쓰세요. ({path})")

    if tool == "Bash":
        cmd = inp.get("command", "")
        for pattern, label in DANGEROUS:
            if re.search(pattern, cmd):
                block(f"{label} 명령은 사람이 직접 실행합니다: {cmd[:120]}")
        hit = touches_protected(cmd)
        if hit and re.search(r"(>|>>|\bmv\b|\bcp\b.*\s" + re.escape(hit) + r"|\btee\b|\bsed\s+-i)", cmd):
            block(f"'{hit}' 폴더에 쓰거나 옮기는 명령입니다: {cmd[:120]}")

    sys.exit(0)


if __name__ == "__main__":
    main()
