#!/usr/bin/env python3
"""소개 페이지 다시 만들기: 저장소 README의 스킬 표를 읽어 template.html에 넣고 index.html을 쓴다.

사용(저장소 최상위 폴더에서):  python docs/promo/make.py
"""
import json
import os
import re

HERE = os.path.dirname(os.path.abspath(__file__))
s = open(os.path.join(HERE, "..", "..", "README.md"), encoding="utf-8").read()
out = []
for icon, name, n, body in re.findall(r"<summary><b>(\S+) (.+?) \((\d+)\)</b></summary>(.*?)</details>", s, re.S):
    for m in re.finditer(r"^\| (🆕 )?\[([a-z0-9-]+)\]\([^)]*\) \| (.+?) \|$", body, re.M):
        out.append({"cat": name, "id": m.group(2), "desc": re.sub(r"🆕 ?", "", m.group(3)).replace("**", ""),
                    "new": bool(m.group(1)) or "🆕" in m.group(3)})
tpl = open(os.path.join(HERE, "template.html"), encoding="utf-8").read()
open(os.path.join(HERE, "index.html"), "w", encoding="utf-8").write(tpl.replace("__SKILLS__", json.dumps(out, ensure_ascii=False)))
print(f"스킬 {len(out)}종 → docs/promo/index.html")
