#!/usr/bin/env python3
"""PostToolUse 훅: 처리 이력을 JSONL로 남기고, 새로 쓴 파일의 SHA-256을 함께 적는다.

이 로그가 '처리 이력 → 지식 노트'의 원천이자, 무엇을 언제 바꿨는지 보여 주는 감사 기록이다.
로그 위치는 RECORDS_LOG (기본 logs/processing_log.jsonl).
"""
import datetime as dt
import hashlib
import json
import os
import sys

LOG = os.environ.get("RECORDS_LOG") or os.path.join(os.environ.get("CLAUDE_PROJECT_DIR", "."), "logs", "processing_log.jsonl")
KST = dt.timezone(dt.timedelta(hours=9))


def sha256(path: str) -> str | None:
    try:
        h = hashlib.sha256()
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(1 << 20), b""):
                h.update(chunk)
        return h.hexdigest()
    except OSError:
        return None


def main() -> None:
    try:
        data = json.load(sys.stdin)
    except json.JSONDecodeError:
        sys.exit(0)
    tool = data.get("tool_name", "")
    inp = data.get("tool_input", {}) or {}
    entry = {
        "time_kst": dt.datetime.now(KST).isoformat(timespec="seconds"),
        "session": data.get("session_id"),
        "tool": tool,
    }
    if tool in ("Write", "Edit", "MultiEdit"):
        path = inp.get("file_path", "")
        entry.update(file=path, sha256=sha256(path))
    elif tool == "Bash":
        entry["command"] = inp.get("command", "")[:500]
    else:
        entry["input_keys"] = sorted(inp.keys())
    os.makedirs(os.path.dirname(LOG) or ".", exist_ok=True)
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(json.dumps(entry, ensure_ascii=False) + "\n")
    sys.exit(0)


if __name__ == "__main__":
    main()
