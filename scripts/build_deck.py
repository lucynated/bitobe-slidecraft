#!/usr/bin/env python3
"""Собирает презентацию из слайдов банка в заданном порядке.

    python3 scripts/build_deck.py 1 152 129 125 181 -o decks/demo.pptx
    python3 scripts/build_deck.py --from decks/2026-11-03_evp.md -o decks/2026-11-03_evp.pptx

Слайды берутся из library/bank.pptx как есть: вёрстка, схемы и медиа не меняются.
Ненужные слайды удаляются, оставшиеся переставляются в нужном порядке.
Один и тот же слайд дважды поставить нельзя (повтор будет пропущен с предупреждением).
С --from номера читаются из таблицы карты слайдов (третья колонка «Слайд»).
"""
import argparse, re, sys
from pathlib import Path
from pptx import Presentation
from pptx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent.parent
BANK = ROOT / "library" / "bank.pptx"


def numbers_from_map(path):
    """Номера слайдов из первой таблицы карты: | № | Тезис | Слайд | ... (номер — в третьей колонке)."""
    nums, in_table = [], False
    for line in Path(path).read_text(encoding="utf-8").splitlines():
        if not line.strip().startswith("|"):
            if in_table:
                break                       # таблица закончилась — дальше пробелы и риски
            continue
        in_table = True
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if len(cells) >= 3 and re.fullmatch(r"\d+", cells[2]):
            nums.append(int(cells[2]))
    return nums


SECTIONS_EXT = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"


def drop_sections(prs):
    """Убрать разделы PowerPoint (боковая панель слайдов), доставшиеся от исходного файла:
    после сборки они пустые или ссылаются на удалённые слайды."""
    ext_lst = prs.part._element.find(qn("p:extLst"))
    if ext_lst is None:
        return
    for ext in list(ext_lst):
        if ext.get("uri") == SECTIONS_EXT:
            ext_lst.remove(ext)


def build(order, out, bank=BANK):
    prs = Presentation(str(bank))
    total = len(prs.slides)
    seen, clean = set(), []
    for n in order:
        if not 1 <= n <= total:
            sys.exit(f"Нет слайда {n}: в банке слайды 1–{total}")
        if n in seen:
            print(f"! слайд {n} уже стоит в колоде, повтор пропущен")
            continue
        seen.add(n); clean.append(n)

    sld_lst = prs.slides._sldIdLst
    ids = list(sld_lst)                      # исходный порядок, индекс = номер-1
    keep = [ids[n - 1] for n in clean]
    for el in ids:                           # убрать всё, что не выбрано
        if el not in keep:
            prs.part.drop_rel(el.rId)
            sld_lst.remove(el)
    for el in keep:                          # переставить в нужном порядке
        sld_lst.remove(el)
    for el in keep:
        sld_lst.append(el)

    drop_sections(prs)
    out = Path(out); out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    print(f"Готово: {out} — {len(clean)} слайдов: {', '.join(map(str, clean))}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("slides", nargs="*", type=int, help="номера слайдов банка по порядку")
    ap.add_argument("--from", dest="map", help="файл карты слайдов (.md)")
    ap.add_argument("-o", "--out", required=True, help="куда сохранить .pptx")
    a = ap.parse_args()
    order = numbers_from_map(a.map) if a.map else a.slides
    if not order:
        sys.exit("Не указаны слайды")
    build(order, a.out)
