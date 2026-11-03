#!/usr/bin/env python3
"""Пересобирает страницу каталога app/index.html и короткий индекс library/index.jsonl
из library/catalog.json и library/examples.json.

    python3 scripts/build_page.py

Запускать после любой правки каталога. Страница открывается двойным кликом, сервер не нужен.
"""
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
LIB = ROOT / "library"
cat = json.loads((LIB / "catalog.json").read_text(encoding="utf-8"))
ex = json.loads((LIB / "examples.json").read_text(encoding="utf-8"))
cards = sorted(cat["slides"], key=lambda c: c["slide"])

# пометки о дублях и вариантах для страницы
dup = {s: g for g in cat.get("duplicates", []) for s in g["slides"]}
for c in cards:
    g = dup.get(c["slide"])
    c["dup_group"] = g["slides"] if g else []
    c["dup_kind"] = g["kind"] if g else ""
    c["keep"] = (c["slide"] == g["keep"]) if g and g.get("keep") else None

# короткий индекс для первого прохода подбора: одна строка на слайд
with open(LIB / "index.jsonl", "w", encoding="utf-8") as f:
    for c in cards:
        extra_dup = c["dup_kind"] == "дубль" and c["keep"] is False
        f.write(json.dumps({
            "slide": c["slide"], "type": c["type"], "role": c["story_role"],
            "practice": c["practice"], "topics": c["topics"], "formats": c["formats"],
            "audience": c["audience"], "client": c["client_data"] != "нет",
            "draft": c["readiness"] == "заготовка", "extra_dup": extra_dup,
            "title": c["title"], "thesis": c["thesis"],
        }, ensure_ascii=False) + "\n")

extra = {"lead": ex["lead"], "dup_lead": ex["dup_lead"], "tests": ex["tests"], "duplicates": cat.get("duplicates", [])}
data = json.dumps({"cards": cards, "extra": extra, "imgBase": "../library/slides/img/"}, ensure_ascii=False).replace("</", "<\\/")
tpl = (ROOT / "app" / "template.html").read_text(encoding="utf-8")
(ROOT / "app" / "index.html").write_text(tpl.replace("/*__DATA__*/null", data), encoding="utf-8")
print(f"Страница: app/index.html, индекс: library/index.jsonl — {len(cards)} слайдов")
