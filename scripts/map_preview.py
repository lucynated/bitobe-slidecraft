#!/usr/bin/env python3
"""Превью карты слайдов: миниатюры выбранных слайдов банка с номером и тезисом.

    python3 scripts/map_preview.py decks/2026-11-03_kp_retail_strat_session.md
    → decks/2026-11-03_kp_retail_strat_session_karta.png

Номер слайда банка сам по себе ничего не говорит человеку, поэтому карту всегда показывают вместе
с этой картинкой: видно, какой слайд стоит под каждым тезисом.
--aspect 1.42 — вписать превью в панель нужной пропорции (для вставки в презентацию).
"""
import argparse, math, os, re, sys
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent.parent
IMG = ROOT / "library" / "slides" / "img"
NAVY, BLUE, GRAY, LIGHT, WHITE = "#19194C", "#182D8D", "#64676A", "#F3F4F6", "#FFFFFF"

FONT_DIRS = [Path.home() / "Library/Fonts", Path("/Library/Fonts"), Path("/System/Library/Fonts/Supplemental"),
             Path("C:/Windows/Fonts"), Path("/usr/share/fonts/truetype/dejavu")]
FONTS = {"regular": ["Onest-Regular.ttf", "Arial.ttf", "arial.ttf", "DejaVuSans.ttf"],
         "medium": ["Onest-Medium.ttf", "Arial Bold.ttf", "arialbd.ttf", "DejaVuSans-Bold.ttf"]}


def font(kind, size):
    for d in FONT_DIRS:
        for name in FONTS[kind]:
            if (d / name).is_file():
                return ImageFont.truetype(str(d / name), size)
    return ImageFont.load_default(size)


def rows_from_map(path):
    """Строки первой таблицы карты: (№, тезис, слайд банка)."""
    out, in_table = [], False
    for line in Path(path).read_text(encoding="utf-8").splitlines():
        if not line.strip().startswith("|"):
            if in_table:
                break
            continue
        in_table = True
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if len(cells) >= 3 and re.fullmatch(r"\d+", cells[2]):
            out.append((cells[0], cells[1], int(cells[2])))
    return out


def wrap(draw, text, fnt, width, max_lines):
    lines, cur = [], ""
    for w in text.split():
        t = (cur + " " + w).strip()
        if draw.textlength(t, font=fnt) <= width:
            cur = t
        else:
            lines.append(cur)
            cur = w
    lines.append(cur)
    if len(lines) > max_lines:
        lines = lines[:max_lines]
        lines[-1] = lines[-1].rstrip(" ,.;:") + "…"
    return lines


def build(map_path, out, cols=None, aspect=None):
    items = rows_from_map(map_path)
    if not items:
        sys.exit(f"В {map_path} нет таблицы карты слайдов")
    n = len(items)
    cols = cols or (3 if n <= 9 else 4)
    rows = math.ceil(n / cols)
    tw, th = 480, 270                       # миниатюра 16:9
    cap_h, gap, pad = 112, 28, 40
    f_num, f_txt = font("medium", 26), font("regular", 24)
    W = pad * 2 + cols * tw + (cols - 1) * gap
    H = pad * 2 + rows * (th + cap_h) + (rows - 1) * gap
    sheet = Image.new("RGB", (W, H), LIGHT)
    d = ImageDraw.Draw(sheet)
    for i, (num, thesis, slide) in enumerate(items):
        x = pad + (i % cols) * (tw + gap)
        y = pad + (i // cols) * (th + cap_h + gap)
        src = IMG / f"s-{slide:03d}.jpg"
        if src.is_file():
            sheet.paste(Image.open(src).convert("RGB").resize((tw, th), Image.LANCZOS), (x, y))
        else:
            d.rectangle([x, y, x + tw, y + th], fill=WHITE)
        d.rectangle([x, y, x + tw - 1, y + th - 1], outline="#D9DBE0")
        d.text((x, y + th + 12), f"{num}  ·  слайд банка {slide}", font=f_num, fill=BLUE)
        for k, line in enumerate(wrap(d, thesis, f_txt, tw, 2)):
            d.text((x, y + th + 48 + k * 30), line, font=f_txt, fill=NAVY)
    if aspect:                              # вписать в панель нужной пропорции, без искажений
        if W / H < aspect:
            pw, ph = round(H * aspect), H
        else:
            pw, ph = W, round(W / aspect)
        panel = Image.new("RGB", (pw, ph), LIGHT)
        panel.paste(sheet, ((pw - W) // 2, (ph - H) // 2))
        sheet = panel
    sheet.save(out, optimize=True)
    print(f"Превью карты: {out} — {n} слайдов")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("map", help="карта слайдов (.md)")
    ap.add_argument("-o", "--out", help="куда сохранить .png (по умолчанию рядом с картой, *_karta.png)")
    ap.add_argument("--cols", type=int, help="миниатюр в ряд")
    ap.add_argument("--aspect", type=float, help="пропорция панели для вставки в презентацию")
    a = ap.parse_args()
    out = a.out or os.path.splitext(a.map)[0] + "_karta.png"
    build(a.map, out, a.cols, a.aspect)
