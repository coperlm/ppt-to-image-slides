#!/usr/bin/env bash
# 用法: ./render-reference.sh <deck.pptx> <out-dir>
# 生成「参考渲染」：LibreOffice headless 转 PDF，再 pdftoppm 拆成逐页 PNG。
# 注意: LibreOffice 的排版 ≠ PowerPoint 真值，它只是一个可复现的一致参考系；
#       真正的验收仍需在 PowerPoint/WPS 里人工比对（见 corpus.md）。
set -euo pipefail

deck="${1:?usage: render-reference.sh <deck.pptx> <out-dir>}"
out="${2:?usage: render-reference.sh <deck.pptx> <out-dir>}"

command -v soffice >/dev/null 2>&1 || { echo "缺少 LibreOffice (soffice)，请先安装" >&2; exit 1; }
command -v pdftoppm >/dev/null 2>&1 || { echo "缺少 poppler (pdftoppm)，请先安装" >&2; exit 1; }

mkdir -p "$out"
tmp="$(mktemp -d)"
trap 'rm -rf "$tmp"' EXIT

soffice --headless --convert-to pdf --outdir "$tmp" "$deck" >/dev/null
pdf="$tmp/$(basename "${deck%.*}").pdf"
[ -f "$pdf" ] || { echo "LibreOffice 没有产出 PDF" >&2; exit 1; }

pdftoppm -png -r 150 "$pdf" "$out/ref"
echo "参考渲染: $(ls "$out"/ref-*.png | wc -l) 页 → $out/ref-*.png"
