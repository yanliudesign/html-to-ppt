---
name: html-to-ppt
description: Convert a self-contained HTML slide deck (1920×1080 slides authored as `.slide` elements) into a `.pptx` file. Two modes — **editable** (每个文本、图片、色块都是原生 PowerPoint 对象，可选中/可编辑) and **raster** (整张幻灯片一张图，视觉 100% 保真但不可编辑). Use when the user says "download as PowerPoint / PPTX", "make an editable version of this deck", "export slides to PowerPoint", or hands over an HTML deck and asks for a `.pptx`. 中文触发词：把 HTML 幻灯片转成 ppt、HTML 转 PPT、网页转 ppt、幻灯片下载成 PowerPoint、下载成 ppt、导出 ppt、导出 PowerPoint、生成可编辑的 ppt、可编辑版本、editable 版本、保留原样的 ppt、整图版 ppt、把这个 deck 转成 ppt、这个网页幻灯片能导出成 ppt 吗。
---

# HTML → PowerPoint

Turn a self-contained HTML slide deck into `.pptx`. Preserves the 1920×1080 stage as a 16:9 slide.

## When to use which mode

Ask the user (or infer from what they say) which they want:

| Mode | Command | Best for | Trade-offs |
|---|---|---|---|
| **Editable** | `scripts/html_to_pptx_editable.py` | Editing text, swapping images, reusing pieces in another deck | Loses gradients / clipped-gradient text / animations. Complex CSS approximated. Slower (Playwright walks DOM per slide). |
| **Raster** | `scripts/html_to_pptx_raster.sh` | Visual fidelity, quick share-outs, print-quality | One flat picture per slide; text is not editable. |

Recommend **editable** by default. Recommend **raster** if the deck relies heavily on gradient text, custom SVG charts, or animation timing that must render exactly as authored.

## Prerequisites

Both modes need macOS (or Linux with substitutions) plus:

- Python 3.9+
- `python-pptx` — `python3 -m pip install --user python-pptx pillow`

**Editable mode** additionally needs:

- `playwright` — `python3 -m pip install --user playwright && python3 -m playwright install chromium`

**Raster mode** additionally needs:

- Any modern Chromium-based browser at the standard `/Applications/` path (Google Chrome, Brave, Arc, Chromium...)
- `pdftoppm` — `brew install poppler`

Check what is already installed before pip-installing anything. Skip the install step if the module imports cleanly.

## Usage

### Editable mode

```bash
python3 scripts/html_to_pptx_editable.py <input.html> <output.pptx>
```

What it does:
1. Launches headless Chromium at 1920×1080.
2. For every `.slide`, force-activates it (sets `.active .visible`, ends reveal animations), then walks the DOM.
3. Emits one native PowerPoint object per visual leaf:
   - Text-leaf blocks → editable text boxes with per-run font-family / size / weight / italic / color / letter-spacing / alignment / line-height / text-transform.
   - `<img>` → picture (WEBP / AVIF are auto-converted to PNG via Pillow; extension-vs-content mismatches are sniffed).
   - Elements with solid `background-color` or borders → rounded/rectangular shape with matching fill + line.
   - Gradient backgrounds, gradient-clipped text, `<svg>` roots → captured as a discrete picture (still individually selectable / replaceable).
4. Assembles a 16:9 PPTX with slide size `1920 × 9525 EMU × 1080 × 9525 EMU` on the blank layout.

### Raster mode

```bash
scripts/html_to_pptx_raster.sh <input.html> <output.pptx>
```

What it does:
1. Injects an `@page{size:1920px 1080px;margin:0}` block plus `@media print` visibility overrides so all slides render (Chrome would otherwise only print the currently-active one).
2. Writes a hidden wrapper file **next to the source** (`.<basename>.print.html`) so relative image paths in the deck still resolve.
3. Prints all slides to a single PDF via Chrome headless.
4. Rasterises each PDF page to 1920×1080 PNG at 96 dpi with `pdftoppm`.
5. Builds a PPTX with one full-bleed picture per slide.

## HTML deck assumptions

The extractor expects the deck convention used across this workspace:

- Each slide is an element with class `slide`, laid out as `position:absolute; inset:0; width:1920px; height:1080px` inside a `.deck-stage` wrapper.
- Slide visibility is toggled by adding/removing `.active` and `.visible` classes.
- Reveal animations use `.reveal` and `.d-stagger > *` starting at `opacity:0`, promoted to visible when the parent `.slide` gets `.visible`.
- Backgrounds may include CSS gradients, SVG roots, and `background-clip: text` gradient titles.

If the deck uses a different scheme, adjust `ACTIVATE_JS` in `html_to_pptx_editable.py` — it is the only slide-aware code.

## Chinese / non-Latin text

- **Editable mode**: font names are copied straight from `computed-style.fontFamily`. Make sure the reader's PowerPoint has the primary CN font (`Noto Sans SC`, `PingFang SC`, `Source Han Sans` …) installed, otherwise PowerPoint falls back to a system CJK font.
- **Raster mode**: fonts are baked into the pixels, so no client-side font install is needed.

## Known caveats (editable mode)

- Elements with `background-image: linear-gradient(...)` are captured as a picture that includes any descendant text. The editable text is then overlaid on top in the correct position — so it looks identical, but if you edit the text the old text is still visible inside the picture underneath. Delete the picture if you need to fully rewrite gradient blocks.
- Inline styling (`<b>`, `<em>`, `<span style>`) is preserved as separate runs inside the same paragraph. Complex flex/grid alignment inside a single text-leaf can drift by 1-2 px.
- Occasional `element is not stable` screenshot timeouts on animated elements (~1 per 50 slides). The build continues; that one decorative bitmap is skipped.

## Where the scripts live

Both scripts are self-contained under this skill's `scripts/` folder:

- `scripts/html_to_pptx_editable.py` (~700 lines) — Playwright + python-pptx.
- `scripts/html_to_pptx_raster.sh` (~70 lines) — Chrome headless + `pdftoppm` + python-pptx.

Copy them into a project's `tools/` directory or run them straight from the skill folder — nothing here is workspace-specific.
