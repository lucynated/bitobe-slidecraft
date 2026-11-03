#!/usr/bin/env bash
# Отрисовывает презентацию в картинки для проверки: render.sh decks/demo.pptx
# Картинки появятся в decks/demo_preview/slide-N.png
set -euo pipefail
f="$1"; dir="$(dirname "$f")"; base="$(basename "${f%.*}")"
out="$dir/${base}_preview"; mkdir -p "$out"
soffice --headless --convert-to pdf --outdir "$out" "$f" >/dev/null 2>&1
pdftoppm -png -r 70 "$out/$base.pdf" "$out/slide"
rm -f "$out/$base.pdf"
ls -1 "$out"
