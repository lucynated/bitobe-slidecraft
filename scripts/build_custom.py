#!/usr/bin/env python3
"""Собирает новую презентацию на шаблоне BITOBE по описанию слайдов.

    python3 scripts/build_custom.py --from decks/2026-11-03_itogi.md -o decks/2026-11-03_itogi.pptx

Основа — library/template.pptx (актуальный шаблон: мастер, макеты, 26 обложек, «Свяжитесь с нами»).
Новые слайды верстаются в макете «Только заголовок» по модульной сетке шаблона: 12 колонок × 5 строк,
каждый блок занимает целое число модулей, содержимое заполняет всю зону контента.
Схемы — фигурами PowerPoint. Слайды банка (тип «банк») копируются из library/bank.pptx как есть,
без заметок докладчика. Формат описания — docs/format_prezentacii.md.
"""
import argparse, copy, math, re, sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part, _Relationship
from pptx.oxml.ns import qn
from pptx.util import Pt

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "library" / "template.pptx"
BANK = ROOT / "library" / "bank.pptx"
BASE = ROOT                     # картинки в описании — относительно файла описания

COVERS = 26                     # слайды 1–26 шаблона — обложки, 27 — «Свяжитесь с нами»
CONTACTS = 27
DEFAULT_COVER = 4

NAVY, BLUE, COPPER = "19194C", "182D8D", "BE9176"
GRAY, LIGHT, RULE, MID = "64676A", "F3F4F6", "E5E7EB", "C8C8C8"
WHITE, BLACK = "FFFFFF", "000000"
BODY, HEAD, SERIF = "Onest", "Onest Medium", "PT Serif"

# Типографическая шкала, pt — одна на все слайды (по образцам шаблона: содержание 28/20/12, текст 14).
# Других размеров в слайдах нет: так текст одного уровня везде выглядит одинаково.
SZ_H2 = 20        # заголовок крупного блока: раздел в содержании
SZ_H3 = 16        # заголовок плашки, пункта, шага
SZ_BODY = 14      # основной текст
SZ_SMALL = 12     # второстепенный текст: подписи, примечания, таблицы, мелкие блоки схем
SZ_LABEL = 10     # надписи прописными, подписи стрелок
SZ_INDEX = 28     # порядковые номера курсивом PT Serif
SZ_BIG = 54       # цифры-факты курсивом PT Serif
SZ_ACCENT = 20    # выводы, цитаты, лид курсивом PT Serif

# Сетка мастера шаблона (направляющие), pt. 12 колонок по 59,5 с промежутком 17,
# 5 строк зоны контента — квадратные модули 59,5. Полоса 108,4–125,4 над первой строкой — отступ от заголовка.
COL_L = [29.25, 105.75, 182.375, 258.875, 335.375, 412.0, 488.5, 565.0, 641.625, 718.125, 794.625, 871.25]
COL_R = [88.75, 165.375, 241.875, 318.375, 395.0, 471.5, 548.0, 624.625, 701.125, 777.625, 854.25, 930.75]
ROW_T = [125.375, 202.0, 278.5, 355.0, 431.625]
ROW_B = [185.0, 261.5, 338.0, 414.625, 491.125]
ROWS = len(ROW_T)
PAD = 14                        # внутренний отступ плашек

WARN = []


def cols(c1, c2):
    return COL_L[c1], COL_R[c2] - COL_L[c1]


def rows(r1, r2):
    return ROW_T[r1], ROW_B[r2] - ROW_T[r1]


def col_spans(n):
    """n равных блоков в ряд по колонкам сетки; None — если 12 колонок на n не делятся."""
    if 12 % n:
        return None
    k = 12 // n
    return [(i * k, i * k + k - 1) for i in range(n)]


def row_spans(r1, r2, n, middle_gap=False):
    """Разбить строки r1..r2 на n полос (большие — первыми)."""
    total = r2 - r1 + 1
    if middle_gap and n == 2 and total % 2 == 1 and total >= 3:
        half = total // 2
        return [(r1, r1 + half - 1), (r2 - half + 1, r2)]
    base, extra = divmod(total, n)
    out, r = [], r1
    for i in range(n):
        size = base + (1 if i < extra else 0)
        out.append((r, r + size - 1))
        r += size
    return out


# ---------- описание слайдов ----------

def parse(path):
    """Описание → (параметры обложки, список слайдов). Формат — docs/format_prezentacii.md."""
    meta, slides, cur = {"заголовок": "", "разделы": []}, [], None
    for raw in Path(path).read_text(encoding="utf-8").splitlines():
        line = raw.rstrip()
        if not line.strip() or line.lstrip().startswith("<!--"):
            continue
        if line.startswith("### "):
            name, _, desc = line[4:].partition(" | ")
            meta["разделы"].append({"название": name.strip(), "описание": desc.strip()})
            cur = None
            continue
        if line.startswith("## "):
            head = line[3:].strip()
            eyebrow, title = ("", head)
            if " | " in head:
                eyebrow, title = (s.strip() for s in head.split(" | ", 1))
            cur = {"надзаголовок": eyebrow, "заголовок": title, "пункты": [],
                   "раздел": len(meta["разделы"]) - 1}
            slides.append(cur)
        elif line.startswith("# ") and cur is None:
            meta["заголовок"] = line[2:].strip()
        elif line.lstrip().startswith("- "):
            if cur is None:
                sys.exit(f"Пункт до первого слайда: {line}")
            cur["пункты"].append(line.lstrip()[2:].strip())
        else:
            m = re.match(r"^([а-яё]+):\s*(.*)$", line.strip(), re.I)
            if not m:
                sys.exit(f"Непонятная строка: {line}")
            key, val = m.group(1).lower(), m.group(2).strip()
            (cur if cur is not None else meta)[key] = val
    return meta, slides


def split_item(item):
    """«Надпись / Заголовок: текст» → (надпись, заголовок, текст)."""
    label, rest = "", item
    head, sep, text = rest.partition(": ")
    if " / " in head:
        label, head = (s.strip() for s in head.split(" / ", 1))
    if not sep:
        return label, "", item.strip()
    return label, head.strip(), text.strip()


def accents(sl):
    return {int(x) for x in re.findall(r"\d+", sl.get("акцент", ""))}


# ---------- оценка высоты текста ----------

def lines_of(text, size, width):
    per_line = max(1, int(width / (size * 0.53)))
    return sum(max(1, math.ceil(len(part) / per_line)) for part in text.split("\n"))


def est_height(paras, width):
    return sum(lines_of(p["text"], p["size"], width) * p["size"] * 1.2 + p["after"] for p in paras)


def check_fit(sl, what, paras, w, h):
    if est_height(paras, w) > h + 2:
        WARN.append(f"«{sl['заголовок'][:40]}…»: {what} может не поместиться — сократите текст")


# ---------- примитивы ----------

def rgb(h):
    return RGBColor.from_string(h)


def P(text, size=SZ_SMALL, color=NAVY, font=BODY, italic=False, after=4, spacing=1.1):
    return {"text": text, "size": size, "color": color, "font": font, "italic": italic,
            "after": after, "spacing": spacing}


def head_p(t, size=SZ_H3, color=NAVY):
    return P(t, size, color, HEAD, after=4)


def body_p(t, size=SZ_BODY, color=GRAY):
    return P(t, size, color, after=3)


def serif_p(t, size=SZ_ACCENT, color=BLACK):
    return P(t, size, color, SERIF, italic=True, spacing=1.15)


def label_p(t, color=BLUE):
    return P(t.upper(), SZ_LABEL, color, HEAD, after=4)


def fill_tf(tf, paras, margin=0, anchor=MSO_ANCHOR.TOP):
    tf.word_wrap = True
    tf.auto_size = None
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Pt(margin)
    tf.vertical_anchor = anchor
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = PP_ALIGN.LEFT
        para.space_after = Pt(p["after"])
        para.line_spacing = p["spacing"]
        run = para.add_run()
        run.text = p["text"]
        f = run.font
        f.size, f.name, f.italic, f.bold = Pt(p["size"]), p["font"], p["italic"], False
        f.color.rgb = rgb(p["color"])


def text(slide, x, y, w, h, paras, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h))
    fill_tf(tb.text_frame, paras, anchor=anchor)
    return tb


def plate(slide, x, y, w, h, fill=LIGHT, border=None, dashed=False):
    """Фирменная плашка: прямоугольник с одним скруглённым углом (правый верхний)."""
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUND_1_RECTANGLE, Pt(x), Pt(y), Pt(w), Pt(h))
    shp.adjustments[0] = min(0.5, 12 / max(min(w, h), 1))
    if fill:
        shp.fill.solid()
        shp.fill.fore_color.rgb = rgb(fill)
    else:
        shp.fill.background()
    if border:
        shp.line.color.rgb = rgb(border)
        shp.line.width = Pt(1.25)
        if dashed:
            shp.line.dash_style = MSO_LINE.DASH
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def rect(slide, x, y, w, h, color):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Pt(x), Pt(y), Pt(w), Pt(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = rgb(color)
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def hline(slide, x, y, w, color=RULE):
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Pt(x), Pt(y), Pt(x + w), Pt(y))
    ln.line.color.rgb = rgb(color)
    ln.line.width = Pt(0.75)


def arrow(slide, x1, y1, x2, y2, elbow=False, dashed=False, color=NAVY):
    c = slide.shapes.add_connector(MSO_CONNECTOR.ELBOW if elbow else MSO_CONNECTOR.STRAIGHT,
                                   Pt(x1), Pt(y1), Pt(x2), Pt(y2))
    c.line.color.rgb = rgb(color)
    c.line.width = Pt(1.5)
    if dashed:
        c.line.dash_style = MSO_LINE.DASH
    tail = etree.SubElement(c.line._get_or_add_ln(), qn("a:tailEnd"))
    tail.set("type", "triangle")
    tail.set("w", "med")
    tail.set("len", "med")


def card(s, sl, x, y, w, h, top, bottom=(), fill=LIGHT, border=None, dashed=False):
    """Плашка: верхний текст прижат к верху, нижний — к низу (как в слайдах банка)."""
    plate(s, x, y, w, h, fill, border, dashed)
    pad = PAD if h > 70 else 8               # плашка в один модуль — отступ меньше
    iw = w - 2 * pad
    th = est_height(top, iw)
    bh = est_height(bottom, iw) if bottom else 0
    check_fit(sl, "текст плашки", list(top) + list(bottom), iw, h - 2 * pad)
    if th + bh + 8 > h - 2 * pad:            # низкая плашка: весь текст одним потоком
        text(s, x + pad, y + pad, iw, h - 2 * pad, list(top) + list(bottom))
        return
    text(s, x + pad, y + pad, iw, max(th, 10), list(top))
    if bottom:
        text(s, x + pad, y + h - pad - bh, iw, bh, list(bottom), anchor=MSO_ANCHOR.BOTTOM)


# ---------- блоки над и под содержимым (по строке сетки) ----------

def block_rows(paras, w, min_rows=1):
    """Сколько строк сетки нужно блоку: 1 модуль или 2 модуля с промежутком."""
    return min_rows if est_height(paras, w) <= 59.5 - 16 else max(min_rows, 2)


def lead_block(s, sl, r1, r2, txt):
    y, h = rows(r1, r2)
    b = rect(s, COL_L[0], y, COL_R[11] - COL_L[0], h, NAVY)
    fill_tf(b.text_frame, [serif_p(txt, color=WHITE)], margin=PAD + 6, anchor=MSO_ANCHOR.MIDDLE)


def quote_block(s, sl, r1, r2, txt):
    y, h = rows(r1, r2)
    paras = [serif_p(txt)]
    check_fit(sl, "цитата или вывод", paras, 901.5, h)
    text(s, COL_L[0], y, 901.5, h, paras, anchor=MSO_ANCHOR.MIDDLE)


def small_text(s, sl, r1, r2, txt):
    y, h = rows(r1, r2)
    _, head, body = split_item(txt)
    hx, hw = cols(0, 2)
    bx, bw = cols(3, 11)
    if head:
        text(s, hx, y, hw, h, [head_p(head)], anchor=MSO_ANCHOR.MIDDLE)
    else:
        bx, bw = cols(0, 11)
    text(s, bx, y, bw, h, [body_p(body)], anchor=MSO_ANCHOR.MIDDLE)


def note_band(s, sl, r1, r2, txt):
    y, h = rows(r1, r2)
    b = rect(s, COL_L[0], y, 901.5, h, LIGHT)
    fill_tf(b.text_frame, [body_p(txt, SZ_SMALL, NAVY)], margin=PAD, anchor=MSO_ANCHOR.MIDDLE)


def titled_band(s, sl, r1, r2, txt, fill):
    """Плашка во всю ширину: заголовок в колонках 0–2, текст в 3–11."""
    y, h = rows(r1, r2)
    _, head, body = split_item(txt)
    plate(s, COL_L[0], y, 901.5, h, fill)
    hx, hw = cols(0, 2)
    bx, bw = cols(3, 11)
    text(s, hx + PAD, y, hw - PAD, h, [head_p(head or "Итог", color=WHITE)], anchor=MSO_ANCHOR.MIDDLE)
    items = [t.strip() for t in body.split(";") if t.strip()]
    paras = [body_p(t, color=WHITE) for t in items]
    check_fit(sl, "плашка", paras, bw - PAD, h - 12)
    text(s, bx, y, bw - PAD, h, paras, anchor=MSO_ANCHOR.MIDDLE)


# ---------- типы слайдов (в строках r1..r2 сетки) ----------

def cards_row(s, sl, r1, r2, kind):
    items, acc = sl["пункты"], accents(sl)
    numbered = sl.get("номера", "").lower() in ("да", "yes")
    n = len(items)
    spans = col_spans(n) if n in (1, 2, 3, 4, 6) else None
    if not spans:
        sys.exit(f"«{sl['заголовок']}»: в типе «{kind}» нужно 1, 2, 3, 4 или 6 пунктов (колонки сетки)")
    y, h = rows(r1, r2)
    for i, (it, (c1, c2)) in enumerate(zip(items, spans), 1):
        x, w = cols(c1, c2)
        label, head, body = split_item(it)
        dark = i in acc
        tc, bc = (WHITE, WHITE) if dark else (NAVY, GRAY)
        top = [P(str(i), SZ_INDEX, WHITE if dark else BLUE, SERIF, italic=True, after=6)] if numbered else []
        top += [label_p(label, WHITE if dark else BLUE)] if label else []
        if kind == "цифры":
            top += [P(head, SZ_BIG, WHITE if dark else BLUE, SERIF, italic=True, after=0)]
            card(s, sl, x, y, w, h, top, [body_p(body, color=tc)], fill=NAVY if dark else LIGHT)
        elif kind == "колонки":
            # заголовок вверху, пункты — по строкам сетки ниже
            plate(s, x, y, w, h, NAVY if dark else LIGHT)
            top += [head_p(head, color=tc)]
            text(s, x + PAD, y + PAD, w - 2 * PAD, est_height(top, w - 2 * PAD), top)
            pts = [t.strip() for t in body.split(";") if t.strip()]
            if r2 > r1 and pts:
                for t, (a, b) in zip(pts, row_spans(r1 + 1, r2, len(pts))):
                    yy, hh = rows(a, b)
                    hline(s, x + PAD, yy, w - 2 * PAD, MID if dark else RULE)
                    text(s, x + PAD, yy + 8, w - 2 * PAD, hh - 8 - PAD / 2, [body_p(t, color=bc)])
                    check_fit(sl, "пункт колонки", [body_p(t)], w - 2 * PAD, hh - 8 - PAD / 2)
        else:
            top += [head_p(head, color=tc)] if head else []
            card(s, sl, x, y, w, h, top, [body_p(body, color=bc)], fill=NAVY if dark else LIGHT)


def list_rows(s, sl, r1, r2):
    items = [split_item(it) for it in sl["пункты"]]
    avail = r2 - r1 + 1
    if len(items) > avail:
        sys.exit(f"«{sl['заголовок']}»: в списке {len(items)} строк, а на слайде место под {avail}")
    hx, hw = cols(0, 3)
    bx, bw = cols(4, 11)
    for i, (_, head, body) in enumerate(items):
        y, h = rows(r1 + i, r1 + i)
        hline(s, COL_L[0], y, 901.5)
        text(s, hx, y + 10, hw, h - 10, [head_p(head)])
        text(s, bx, y + 10, bw, h - 10, [body_p(body)])
        check_fit(sl, "строка списка", [body_p(body)], bw, h - 10)


def checklist(s, sl, r1, r2):
    items = sl["пункты"]
    k = math.ceil(len(items) / 2)
    avail = r2 - r1 + 1
    if k > avail:
        sys.exit(f"«{sl['заголовок']}»: в чек-листе {len(items)} пунктов, а места под {avail * 2}")
    picks = [r1 + round(i * (avail - 1) / (k - 1)) for i in range(k)] if k > 1 else [r1]
    for i, t in enumerate(items):
        c1, c2 = (0, 5) if i % 2 == 0 else (6, 11)
        x, w = cols(c1, c2)
        y, h = rows(picks[i // 2], picks[i // 2])
        rect(s, x, y + 5, 10, 10, COPPER)
        text(s, x + 26, y, w - 26 - 17, h, [body_p(t, color=NAVY)])
        check_fit(sl, "пункт чек-листа", [body_p(t)], w - 43, h)


def steps(s, sl, r1, r2):
    items = sl["пункты"]
    spans = col_spans(len(items))
    if not spans:
        sys.exit(f"«{sl['заголовок']}»: в шагах нужно 2, 3, 4 или 6 пунктов (колонки сетки)")
    y, h = rows(r1, r2)
    for i, (it, (c1, c2)) in enumerate(zip(items, spans)):
        x, w = cols(c1, c2)
        _, head, body = split_item(it)
        text(s, x, y, w, 40, [P(str(i + 1), SZ_INDEX, BLUE, SERIF, italic=True, after=0)])
        paras = ([head_p(head)] if head else []) + [body_p(body, color=NAVY)]
        ty = y + 44
        text(s, x, ty, w - 8, y + h - ty, paras)
        check_fit(sl, "шаг", paras, w - 8, y + h - ty)


def chain(s, sl, r1, r2):
    """Цепочка шагов: до 6 — в строку, 7–12 — «змейкой» в две строки."""
    items, acc = sl["пункты"], accents(sl)
    n = len(items)
    if not 2 <= n <= 12:
        sys.exit(f"«{sl['заголовок']}»: в цепочке нужно 2–12 шагов")
    per_row = n if n <= 6 else math.ceil(n / 2)
    spans = col_spans(per_row)
    if spans:
        xs = [cols(c1, c2) for c1, c2 in spans]
    else:                                # 5 шагов в строку: равные блоки по полям сетки, стрелки в промежутках
        gap = 34
        w = (901.5 - gap * (per_row - 1)) / per_row
        xs = [(COL_L[0] + i * (w + gap), w) for i in range(per_row)]
    band_rows = row_spans(r1, r2, 1 if n <= per_row else 2, middle_gap=True)
    boxes = []
    for i in range(n):
        r, c = divmod(i, per_row)
        if r == 1:
            c = per_row - 1 - c          # вторая строка идёт справа налево
        x, w = xs[c]
        y, h = rows(*band_rows[r])
        boxes.append((x, y, w, h))
    for i, (it, (x, y, w, h)) in enumerate(zip(items, boxes), 1):
        _, head, body = split_item(it)
        dark = i in acc
        last = i == n and not dark
        num = P(str(i), SZ_INDEX, WHITE if dark else BLUE, SERIF, italic=True, after=6) if h > 150 \
            else P(f"{i:02d}", SZ_LABEL, WHITE if dark else BLUE, HEAD, after=4)
        top = [num, head_p(head or "", color=WHITE if dark else NAVY)]
        bottom = [body_p(body, SZ_SMALL, WHITE if dark else GRAY)]
        card(s, sl, x, y, w, h, top, bottom, fill=NAVY if dark else (WHITE if last else LIGHT),
             border=COPPER if last else None)
    for i in range(n - 1):
        (x1, y1, w1, h1), (x2, y2, w2, h2) = boxes[i], boxes[i + 1]
        if y1 == y2 and x2 > x1:
            arrow(s, x1 + w1, y1 + h1 / 2, x2, y2 + h2 / 2)
        elif y1 == y2:
            arrow(s, x1, y1 + h1 / 2, x2 + w2, y2 + h2 / 2)
        else:
            arrow(s, x1 + w1 / 2, y1 + h1, x2 + w2 / 2, y2)


SCHEME_COLS = {"A": (0, 2), "B": (4, 7), "C": (9, 11)}   # колонки 3 и 8 — под стрелки и подписи


def scheme(s, sl, r1, r2):
    """Блоки в трёх колонках сетки (A 0–2, B 4–7, C 9–11), строки — модули сетки.
    Пункт: «A1 Заголовок: текст», «B1-3 …» — блок на строки 1–3, «!C2 …» — тёмный, «?A3 …» — пунктирный.
    связи: «A1->B1 подпись; A2-->B1» (--> — пунктир)."""
    boxes = {}
    for it in sl["пункты"]:
        m = re.match(r"^([!?]?)([ABC])(\d)(?:-(\d))?\s+(.*)$", it)
        if not m:
            sys.exit(f"«{sl['заголовок']}»: блок схемы «{it}» должен начинаться с A1, B2-3, !C1 …")
        mark, col, a, b, rest = m.groups()
        boxes[f"{col}{a}"] = dict(mark=mark, col=col, a=int(a), b=int(b or a), rest=rest)
    nrows = max(b["b"] for b in boxes.values())
    bands = row_spans(r1, r2, nrows)
    for key, b in boxes.items():
        b["x"], b["w"] = cols(*SCHEME_COLS[b["col"]])
        b["y"] = ROW_T[bands[b["a"] - 1][0]]
        b["h"] = ROW_B[bands[b["b"] - 1][1]] - b["y"]
        _, head, body = split_item(b["rest"])
        dark = b["mark"] == "!"
        top = [head_p(head or body, color=WHITE if dark else NAVY)]
        bottom = []
        if head:
            for seg in [t.strip() for t in body.split(" // ") if t.strip()]:
                if seg.startswith("`") and seg.endswith("`"):
                    top.append(body_p(seg.strip("`"), SZ_SMALL, WHITE if dark else BLUE))
                else:
                    bottom.append(body_p(seg, SZ_SMALL, WHITE if dark else GRAY))
        card(s, sl, b["x"], b["y"], b["w"], b["h"], top, bottom,
             fill=NAVY if dark else (WHITE if b["mark"] == "?" else LIGHT),
             border=MID if b["mark"] == "?" else None, dashed=b["mark"] == "?")
    for link in [t.strip() for t in sl.get("связи", "").split(";") if t.strip()]:
        m = re.match(r"^([ABC]\d)\s*(-->|->)\s*([ABC]\d)\s*(.*)$", link)
        if not m or m.group(1) not in boxes or m.group(3) not in boxes:
            sys.exit(f"«{sl['заголовок']}»: непонятная связь «{link}»")
        a, kind, z, label = boxes[m.group(1)], m.group(2), boxes[m.group(3)], m.group(4).strip()
        dashed = kind == "-->"
        if a["col"] != z["col"]:
            lo, hi = max(a["y"], z["y"]), min(a["y"] + a["h"], z["y"] + z["h"])
            if hi - lo > 20:
                y1 = y2 = (lo + hi) / 2
            else:
                y1, y2 = a["y"] + a["h"] / 2, z["y"] + z["h"] / 2
            x1, x2 = (a["x"] + a["w"], z["x"]) if a["x"] < z["x"] else (a["x"], z["x"] + z["w"])
            arrow(s, x1, y1, x2, y2, elbow=y1 != y2, dashed=dashed)
            if label:
                text(s, min(x1, x2) + 4, y1 - 17, abs(x2 - x1) - 8, 15, [body_p(label, SZ_LABEL)])
        else:
            x = a["x"] + a["w"] / 2
            y1, y2 = (a["y"] + a["h"], z["y"]) if a["y"] < z["y"] else (a["y"], z["y"] + z["h"])
            arrow(s, x, y1, x, y2, dashed=dashed)
            if label:
                text(s, x + 8, (y1 + y2) / 2 - 8, 120, 15, [body_p(label, SZ_LABEL)])


# ---------- картинки и таблицы ----------

def image_path(sl, key="файл"):
    f = (BASE / sl.get(key, "")).resolve()
    if not f.is_file():
        sys.exit(f"«{sl['заголовок']}»: нет картинки {f}")
    return f


def place_image(s, sl, path, x, y, w, h):
    """Картинка заполняет ячейку сетки целиком: пропорции сохраняются, лишнее срезается (по центру
    по ширине, снизу по высоте). Срез больше 5% — предупреждение: картинку надо делать в пропорции ячейки."""
    from PIL import Image
    iw, ih = Image.open(path).size
    pic = s.shapes.add_picture(str(path), Pt(x), Pt(y), Pt(w), Pt(h))
    img_ar, box_ar = iw / ih, w / h
    if img_ar > box_ar:
        cut = 1 - box_ar / img_ar
        pic.crop_left = pic.crop_right = cut / 2
    else:
        cut = 1 - img_ar / box_ar
        pic.crop_bottom = cut
    if cut > 0.05:
        WARN.append(f"«{sl['заголовок'][:40]}…»: картинка {Path(path).name} обрезана на {cut:.0%} — "
                    f"сделайте её в пропорции {w / h:.2f}")
    pic.line.color.rgb = rgb(RULE)
    pic.line.width = Pt(0.75)


def image_with_text(s, sl, r1, r2):
    """Картинка на части колонок, пункты с другой стороны — по строкам сетки, с линейками."""
    wide = int(sl.get("ширина", "7"))
    left = sl.get("сторона", "слева").lower().startswith("сл")
    img_cols = (0, wide - 1) if left else (12 - wide, 11)
    txt_cols = (wide, 11) if left else (0, 11 - wide)
    x, w = cols(*img_cols)
    y, h = rows(r1, r2)
    place_image(s, sl, image_path(sl), x, y, w, h)
    tx, tw = cols(*txt_cols)
    items = [split_item(it) for it in sl["пункты"]]
    if not items:
        return
    # пункты начинаются на линиях сетки с равным шагом: 2 пункта — строки 1 и 3, 3 — 1, 3, 5, 4–5 — подряд
    total = r2 - r1 + 1
    n = len(items)
    if n > total:
        sys.exit(f"«{sl['заголовок']}»: рядом с картинкой не больше {total} пунктов")
    step = (total - 1) // (n - 1) if n > 1 else total
    for i, (label, head, body) in enumerate(items):
        a = r1 + i * step
        b = r2 if i == n - 1 else a + step - 1
        yy = ROW_T[a]
        hh = (ROW_B[b] if b == r2 else ROW_T[b + 1]) - yy      # до следующего пункта, с промежутком
        hline(s, tx, yy, tw)
        paras = ([label_p(label)] if label else []) + ([head_p(head)] if head else []) + [body_p(body)]
        text(s, tx, yy + 8, tw, hh - 8, paras)
        check_fit(sl, "пункт рядом с картинкой", paras, tw, hh - 8)


def images_row(s, sl, r1, r2):
    """2–4 картинки в ряд с подписями: картинки в строках r1..r2-1, подписи — в последней строке."""
    items = [split_item(it) for it in sl["пункты"]]
    spans = col_spans(len(items))
    if not spans:
        sys.exit(f"«{sl['заголовок']}»: картинок в ряд — 1, 2, 3 или 4")
    y, h = rows(r1, r2 - 1)
    cy, ch = rows(r2, r2)
    for (label, head, body), (c1, c2) in zip(items, spans):
        x, w = cols(c1, c2)
        f = (BASE / head).resolve()
        if not f.is_file():
            sys.exit(f"«{sl['заголовок']}»: нет картинки {f}")
        place_image(s, sl, f, x, y, w, h)
        paras = ([label_p(label)] if label else []) + [body_p(body, color=NAVY)]
        text(s, x, cy, w, ch, paras)
        check_fit(sl, "подпись картинки", paras, w, ch)


def table(s, sl, r1, r2):
    """Таблица на всю ширину сетки: первая строка — шапка, ячейки через « | ».
    Без линеек между строками (стандарт BITOBE): строки разделяет воздух."""
    data = [[c.strip() for c in it.split("|")] for it in sl["пункты"]]
    ncol = len(data[0])
    weights = [float(x) for x in re.findall(r"[\d.]+", sl.get("ширины", ""))] or [1] * ncol
    x, w = cols(0, 11)
    y, h = rows(r1, r2)
    gf = s.shapes.add_table(len(data), ncol, Pt(x), Pt(y), Pt(w), Pt(h))
    tbl = gf.table
    style = tbl._tbl.tblPr.find(qn("a:tableStyleId"))
    if style is None:
        style = etree.SubElement(tbl._tbl.tblPr, qn("a:tableStyleId"))
    style.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"      # «Без стиля, без сетки»
    tbl.first_row = False
    tbl.horz_banding = False
    for j, wt in enumerate(weights):
        tbl.columns[j].width = Pt(w * wt / sum(weights))
    row_h = h / len(data)
    for i, row in enumerate(data):
        tbl.rows[i].height = Pt(row_h)
        for j in range(ncol):
            cell = tbl.cell(i, j)
            cell.fill.background()
            cell.margin_left, cell.margin_right = Pt(0), Pt(10)
            cell.margin_top = cell.margin_bottom = Pt(3)
            val = row[j] if j < len(row) else ""
            para = label_p(val) if i == 0 else body_p(val, SZ_SMALL, NAVY if j < 2 else GRAY)
            fill_tf(cell.text_frame, [para], anchor=MSO_ANCHOR.TOP)


# ---------- слайды ----------

def layout(prs, name):
    for l in prs.slide_layouts:
        if l.name.strip() == name:
            return l
    sys.exit(f"В шаблоне нет макета «{name}»")


def progress_bar(ph, sections, current):
    """Навигация шаблона «Прогресс бар»: все разделы в строке над заголовком, текущий — жирным."""
    tf = ph.text_frame
    para = tf.paragraphs[0]
    for i, sec in enumerate(sections):
        if i:
            para.add_run().text = "  |  "
        r = para.add_run()
        r.text = sec["название"]
        if i == current:
            r.font.bold = True
            r.font.color.rgb = rgb(NAVY)


def content_slide(prs, sl, sections):
    s = prs.slides.add_slide(layout(prs, "Только заголовок"))
    s.shapes.title.text = sl["заголовок"]
    for ph in list(s.placeholders):
        idx = ph.placeholder_format.idx
        if idx == 11:
            if sections and sl["раздел"] >= 0:
                progress_bar(ph, sections, sl["раздел"])
            elif sl["надзаголовок"]:
                ph.text = sl["надзаголовок"]
            else:
                ph._element.getparent().remove(ph._element)
        elif idx == 10:
            ph._element.getparent().remove(ph._element)
    return s


def contents_slide(prs, sections, title):
    """Содержание по образцу шаблона: карточка на раздел — номер, название, описание, «Перейти →»."""
    spans = col_spans(len(sections))
    if not spans:
        sys.exit("Разделов для содержания нужно 2, 3, 4 или 6 — по колонкам сетки")
    s = prs.slides.add_slide(layout(prs, "Только заголовок"))
    s.shapes.title.text = title
    for ph in list(s.placeholders):
        if ph.placeholder_format.idx in (10, 11):
            ph._element.getparent().remove(ph._element)
    y, h = rows(0, ROWS - 1)
    buttons = []
    for i, (sec, (c1, c2)) in enumerate(zip(sections, spans), 1):
        x, w = cols(c1, c2)
        plate(s, x, y, w, h, LIGHT)
        iw = w - 2 * PAD
        top = [P(f"{i:02d}", SZ_INDEX, BLUE, SERIF, italic=True, after=6), head_p(sec["название"], SZ_H2)]
        th = est_height(top, iw)
        text(s, x + PAD, y + PAD, iw, th, top)
        hline(s, x + PAD, y + PAD + th + 6, iw, MID)
        dh = est_height([body_p(sec["описание"], SZ_SMALL)], iw)
        text(s, x + PAD, y + PAD + th + 14, iw, dh, [body_p(sec["описание"], SZ_SMALL)])
        sec["_list"] = text(s, x + PAD, ROW_T[2], iw, ROW_B[ROWS - 1] - ROW_T[2] - 40, [])
        btn = plate(s, x + PAD, y + h - PAD - 24, 96, 24, COPPER)
        fill_tf(btn.text_frame, [P("Перейти →", SZ_SMALL, WHITE, after=0)], margin=4, anchor=MSO_ANCHOR.MIDDLE)
        btn.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        buttons.append(btn)
    return s, buttons


def fill_contents(sections):
    """Список слайдов в карточках содержания: номер страницы и короткое название."""
    for sec in sections:
        tb = sec.get("_list")
        if tb is None:
            continue
        paras = []
        for page, name in sec.get("_slides", []):
            paras.append(P(f"{page:02d}   {name}", SZ_SMALL, NAVY, after=8))
        tb.text_frame.paragraphs[0].text = ""
        fill_tf(tb.text_frame, paras)


def divider_slide(prs, sl):
    s = prs.slides.add_slide(layout(prs, "Разделитель единичный"))
    s.shapes.title.text = sl["заголовок"].replace(" // ", "\v")
    for ph in s.placeholders:
        if ph.placeholder_format.idx == 10:
            ph.text = sl.get("текст", "")
    return s


def clone_part(src_part, parent, reltype, rId=None):
    """Копия части пакета (диаграмма, картинка, вложенная таблица) со всеми её связями."""
    pkg = parent.package
    tmpl = re.sub(r"\d*(\.\w+)$", r"%d\1", str(src_part.partname))
    new = Part(pkg.next_partname(tmpl), src_part.content_type, pkg, src_part.blob)
    if rId is None:
        rId = parent.rels._add_relationship(reltype, new)
    else:
        parent.rels._rels[rId] = _Relationship(parent.rels._base_uri, rId, reltype, "Internal", new)
    for child_rId, rel in src_part.rels.items():
        if rel.is_external:
            new.rels._rels[child_rId] = _Relationship(new.rels._base_uri, child_rId, rel.reltype,
                                                      "External", rel.target_ref)
        else:
            clone_part(rel.target_part, new, rel.reltype, child_rId)
    return rId


def bank_slide(prs, bank, num, used):
    """Слайд банка как есть: фигуры, картинки, диаграммы. Заметки докладчика не переносятся."""
    if not 1 <= num <= len(bank.slides):
        sys.exit(f"Нет слайда {num} в банке")
    if num in used:
        WARN.append(f"слайд банка {num} стоит дважды")
    used.add(num)
    src = bank.slides[num - 1]
    s = prs.slides.add_slide(layout(prs, src.slide_layout.name.strip()))
    tree = s.shapes._spTree
    for el in list(tree)[2:]:
        tree.remove(el)
    bg = src._element.cSld.bg
    if bg is not None:
        s._element.cSld.insert(0, copy.deepcopy(bg))
    rid_map = {}
    for rId, rel in src.part.rels.items():
        if rel.reltype in (RT.SLIDE_LAYOUT, RT.NOTES_SLIDE):
            continue
        if rel.is_external:
            rid_map[rId] = s.part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        else:
            rid_map[rId] = clone_part(rel.target_part, s.part, rel.reltype)
    for el in list(src.shapes._spTree)[2:]:
        new_el = copy.deepcopy(el)
        for node in new_el.iter():
            for attr, val in list(node.attrib.items()):
                if attr.startswith("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}") \
                        and val in rid_map:
                    node.set(attr, rid_map[val])
        tree.append(new_el)
    return s


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


def fill_cover(slide, meta):
    for ph in slide.placeholders:
        idx = ph.placeholder_format.idx
        if idx == 0:
            ph.text = meta.get("заголовок", "").replace(" // ", "\v")
        elif idx == 15:
            ph.text = meta.get("дата", "")
        elif idx == 16:
            ph.text = meta.get("подзаголовок", "")
        elif idx == 12:
            who = meta.get("для", "")
            ph.text = f"Подготовлено для {who}" if who else ""


TYPES = {"карточки", "цифры", "колонки", "список", "чеклист", "шаги", "цепочка", "схема", "раздел", "банк",
         "картинка", "картинки", "таблица"}
ABOVE = ("лид", "цитата")
BELOW = ("вывод", "примечание", "пример", "плашка", "текст")


def layout_content(s, sl, kind):
    """Распределить 5 строк сетки: блоки сверху, содержимое, блоки снизу."""
    plan = []
    for key in ABOVE + BELOW:
        if not sl.get(key):
            continue
        val = sl[key]
        if key in ("пример", "плашка"):
            _, head, body = split_item(val)
            items = [t for t in body.split(";") if t.strip()]
            need = 2 if len(items) > 2 or len(body) > 90 else 1
        elif key == "лид":
            need = block_rows([serif_p(val)], 860)
        elif key in ("цитата", "вывод"):
            need = block_rows([serif_p(val)], 901)
        else:
            need = block_rows([body_p(val)], 640)
        plan.append((key, need))
    used = sum(n for _, n in plan)
    main = ROWS - used
    if kind == "список":
        main = max(main, len(sl["пункты"]))
    if main < 1:
        sys.exit(f"«{sl['заголовок']}»: слишком много блоков для одного слайда — разделите его")
    r = 0
    for key, need in [p for p in plan if p[0] in ABOVE]:
        draw_block(s, sl, key, r, r + need - 1)
        r += need
    m1, m2 = r, r + main - 1
    if kind in ("карточки", "цифры", "колонки"):
        cards_row(s, sl, m1, m2, kind)
    elif kind == "список":
        list_rows(s, sl, m1, m2)
    elif kind == "чеклист":
        checklist(s, sl, m1, m2)
    elif kind == "шаги":
        steps(s, sl, m1, m2)
    elif kind == "цепочка":
        chain(s, sl, m1, m2)
    elif kind == "схема":
        scheme(s, sl, m1, m2)
    elif kind == "картинка":
        image_with_text(s, sl, m1, m2)
    elif kind == "картинки":
        images_row(s, sl, m1, m2)
    elif kind == "таблица":
        table(s, sl, m1, m2)
    r = m2 + 1
    for key, need in [p for p in plan if p[0] in BELOW]:
        if r + need - 1 >= ROWS:
            sys.exit(f"«{sl['заголовок']}»: не хватает строк сетки — сократите список или блоки")
        draw_block(s, sl, key, r, r + need - 1)
        r += need


def draw_block(s, sl, key, r1, r2):
    val = sl[key]
    if key == "лид":
        lead_block(s, sl, r1, r2, val)
    elif key == "цитата":
        quote_block(s, sl, r1, r2, val)
    elif key == "вывод":
        quote_block(s, sl, r1, r2, val)
    elif key == "текст":
        small_text(s, sl, r1, r2, val)
    elif key == "примечание":
        note_band(s, sl, r1, r2, val)
    elif key == "пример":
        titled_band(s, sl, r1, r2, val, BLUE)
    elif key == "плашка":
        titled_band(s, sl, r1, r2, val, NAVY)


def build(spec, out):
    global BASE
    BASE = Path(spec).resolve().parent
    meta, slides = parse(spec)
    prs = Presentation(str(TEMPLATE))
    bank = None
    used = set()

    lst = prs.slides._sldIdLst
    ids = list(lst)
    cover_n = int(meta.get("обложка", DEFAULT_COVER))
    if not 1 <= cover_n <= COVERS:
        sys.exit(f"Обложки 1–{COVERS}, указана {cover_n}")
    cover_el, contacts_el = ids[cover_n - 1], ids[CONTACTS - 1]
    for el in ids:
        if el not in (cover_el, contacts_el):
            prs.part.drop_rel(el.rId)
            lst.remove(el)
    prs.part.rename_slide_parts([el.rId for el in lst])   # иначе новые слайды займут имена оставшихся
    fill_cover(prs.slides[0], meta)
    if meta.get("заметки"):
        prs.slides[0].notes_slide.notes_text_frame.text = meta["заметки"]

    sections = meta["разделы"]
    buttons, first_of = [], {}
    if len(sections) >= 2 and meta.get("содержание", "да").lower() not in ("нет", "no"):
        _, buttons = contents_slide(prs, sections, meta.get("содержание_заголовок", "Содержание"))

    for n, sl in enumerate(slides, 2):
        kind = sl.get("тип", "карточки").lower()
        if kind not in TYPES:
            sys.exit(f"Слайд {n}: неизвестный тип «{kind}». Есть: {', '.join(sorted(TYPES))}")
        if kind == "банк":
            bank = bank or Presentation(str(BANK))
            s = bank_slide(prs, bank, int(sl.get("номер", "0")), used)
        elif kind == "раздел":
            s = divider_slide(prs, sl)
        else:
            s = content_slide(prs, sl, sections)
            layout_content(s, sl, kind)
        if sl["раздел"] >= 0:
            first_of.setdefault(sl["раздел"], s)
            short = sl["надзаголовок"] or " ".join(sl["заголовок"].split()[:4])
            sections[sl["раздел"]].setdefault("_slides", []).append((len(prs.slides) - 1, short))  # −1: контакты пока стоят вторыми
        if sl.get("заметки"):
            s.notes_slide.notes_text_frame.text = sl["заметки"]

    fill_contents(sections)
    for i, btn in enumerate(buttons):
        if i in first_of:
            btn.click_action.target_slide = first_of[i]

    if meta.get("контакты", "да").lower() in ("нет", "no"):
        prs.part.drop_rel(contacts_el.rId)
        lst.remove(contacts_el)
    else:
        lst.remove(contacts_el)
        lst.append(contacts_el)
    prs.part.rename_slide_parts([el.rId for el in lst])
    drop_sections(prs)

    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    print(f"Готово: {out} — {len(prs.slides)} слайдов")
    for w in WARN:
        print("! " + w)


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--from", dest="spec", required=True, help="описание презентации (.md)")
    ap.add_argument("-o", "--out", required=True, help="куда сохранить .pptx")
    a = ap.parse_args()
    build(a.spec, a.out)
