#!/usr/bin/env python3
"""Проверяет library/catalog.json по схеме taxonomy.md: поля, допустимые значения, ссылки.

    python3 scripts/check_catalog.py

Код выхода 1, если есть ошибки. Запускать перед каждым коммитом каталога.
"""
import json, sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TYPE = {"служебный", "обложка", "о BITOBE", "люди", "продукт", "почему важно", "подход", "модель", "кейс",
        "исследование", "проект", "план и бюджет", "интерактив", "макет"}
ROLE = {"открытие", "проблема", "решение", "доказательство", "детали", "следующий шаг", "служебный"}
PRACTICE = {"Управление персоналом и оргразвитие", "Оценка персонала", "Командные трансформации",
            "Обучение и развитие", "Визуальный и коммуникационный консалтинг", "BITOBE в целом"}
TOPICS = set("""корпоративная культура|ценности|EVP и бренд работодателя|HR-стратегия и HR-функция|оргструктура и роли|
оценка руководителей|управление талантами и резерв|мотивация и вознаграждение|исследования персонала|
командная эффективность|управленческая команда|лидерство|целеполагание|управление изменениями|
проектное управление|операционная эффективность|клиентский опыт и сервис|обучение и программы развития|
бизнес-сессии и игры|креативность и инновации|ИИ|внутренние коммуникации|о компании""".replace("\n", "").split("|"))
FORMATS = {"КП и продажа", "стратегическая сессия / воркшоп", "выступление / лекция", "отчёт по проекту", "презентация компании"}
AUDIENCE = {"собственники и CEO", "HR-директора", "руководители", "участники сессии", "широкая аудитория"}
FIELDS = ["slide", "title", "thesis", "use_when", "avoid_when", "type", "story_role", "practice", "topics", "formats",
          "audience", "client_data", "readiness", "layout_reusable", "detail", "related", "note"]

cat = json.loads((ROOT / "library" / "catalog.json").read_text(encoding="utf-8"))
nums = {c["slide"] for c in cat["slides"]}
errs = []
for c in cat["slides"]:
    n = c.get("slide", "?")
    miss = [f for f in FIELDS if f not in c]
    if miss: errs.append(f"{n}: нет полей {miss}")
    if c.get("type") not in TYPE: errs.append(f"{n}: type «{c.get('type')}»")
    if c.get("story_role") not in ROLE: errs.append(f"{n}: story_role «{c.get('story_role')}»")
    for field, allowed in (("practice", PRACTICE), ("topics", TOPICS), ("formats", FORMATS), ("audience", AUDIENCE)):
        bad = [v for v in c.get(field, []) if v not in allowed]
        if bad: errs.append(f"{n}: {field} {bad}")
    if c.get("readiness") not in ("готов", "заготовка"): errs.append(f"{n}: readiness")
    if not (c.get("client_data") == "нет" or str(c.get("client_data", "")).startswith("клиент:")): errs.append(f"{n}: client_data")
    if len(c.get("thesis", "")) < 20: errs.append(f"{n}: слишком короткий thesis")
    bad_rel = [r for r in c.get("related", []) if r not in nums]
    if bad_rel: errs.append(f"{n}: related на несуществующие {bad_rel}")
for g in cat.get("duplicates", []):
    if any(s not in nums for s in g["slides"]): errs.append(f"дубли {g['slides']}: несуществующий слайд")

print("\n".join(errs) if errs else f"Каталог в порядке: {len(nums)} карточек")
sys.exit(1 if errs else 0)
