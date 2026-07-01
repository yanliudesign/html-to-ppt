#!/usr/bin/env bash
# Convert a self-contained HTML slide deck (1920x1080 slides) to PPTX.
# Usage: html_to_pptx.sh <input.html> <output.pptx>
set -euo pipefail

IN="$1"; OUT="$2"
[[ -f "$IN" ]] || { echo "input not found: $IN"; exit 1; }

CHROME="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
[[ -x "$CHROME" ]] || CHROME="/Applications/Chromium.app/Contents/MacOS/Chromium"
[[ -x "$CHROME" ]] || CHROME="/Applications/Brave Browser.app/Contents/MacOS/Brave Browser"

TMP=$(mktemp -d)
trap 'rm -rf "$TMP"; rm -f "$WRAP"' EXIT

BASENAME=$(basename "$IN" .html)
# Wrap file must live next to the source so relative screenshot paths still resolve.
SRC_DIR=$(cd "$(dirname "$IN")" && pwd)
WRAP="$SRC_DIR/.${BASENAME}.print.html"
PDF="$TMP/${BASENAME}.pdf"
IMG_PREFIX="$TMP/slide"

# Inject @page rule so Chrome prints each 1920x1080 slide as one full page.
# We insert it right before the closing </style> tag of the deck.
python3 - "$IN" "$WRAP" <<'PY'
import sys, pathlib, re
src = pathlib.Path(sys.argv[1]).read_text()
# Force every slide + every reveal-animation element to be fully visible in print,
# and set the page size to match the 1920x1080 slide dimensions.
inject = (
  "@page{size:1920px 1080px;margin:0}"
  "@media print{"
    ".slide,.slide *{visibility:visible!important;opacity:1!important;transform:none!important;animation:none!important;transition:none!important}"
    ".reveal,.d-stagger>*{opacity:1!important;transform:none!important}"
  "}"
)
out = re.sub(r'</style>', inject + '</style>', src, count=1)
pathlib.Path(sys.argv[2]).write_text(out)
PY

echo "→ printing to PDF via Chrome…"
"$CHROME" --headless=new --disable-gpu --no-pdf-header-footer \
  --print-to-pdf-no-header \
  --print-to-pdf="$PDF" \
  "file://$WRAP" 2>/dev/null

echo "→ rasterising PDF pages @ 96dpi (→ 1920x1080 px)…"
pdftoppm -r 96 -png "$PDF" "$IMG_PREFIX"

echo "→ building PPTX…"
python3 - "$OUT" "$TMP" <<'PY'
import sys, glob, os
from pptx import Presentation
from pptx.util import Emu

out_path, tmp = sys.argv[1], sys.argv[2]
imgs = sorted(glob.glob(os.path.join(tmp, "slide-*.png")))
assert imgs, "no slide images produced"

prs = Presentation()
# 16:9 at 1920x1080; use exact EMUs (1 px = 9525 EMU at 96 dpi baseline)
prs.slide_width  = Emu(1920 * 9525)
prs.slide_height = Emu(1080 * 9525)
blank = prs.slide_layouts[6]  # blank layout

for img in imgs:
    slide = prs.slides.add_slide(blank)
    slide.shapes.add_picture(img, 0, 0, width=prs.slide_width, height=prs.slide_height)

prs.save(out_path)
print(f"✓ wrote {out_path}  ({len(imgs)} slides)")
PY
