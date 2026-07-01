# html-to-ppt

Convert self-contained HTML slide decks (1920 × 1080 slides authored as `.slide` elements) to `.pptx`.

Two modes:

| Mode | Result |
|---|---|
| **Editable** | Every text run, image, and colored box becomes a **native PowerPoint object** you can select, restyle, and edit. |
| **Raster** | One full-bleed image per slide. Fast, always visually faithful, but not editable. |

Written for HTML decks that use a `1920×1080` fixed stage — the same convention as the [frontend-slides](../frontend-slides) skill and most single-file HTML presentations.

---

## 中文说明 (Chinese)

把一份自包含的 HTML 幻灯片（1920×1080，每页是一个 `.slide` 元素）转成 `.pptx`。两种模式：

| 模式 | 效果 |
|---|---|
| **可编辑 Editable** | 每段文字、每张图片、每个色块都变成 **原生 PowerPoint 对象**，可以选中、改字、换图、复用到别的 deck。 |
| **整图 Raster** | 每页导出为一张整图，视觉 100% 保真，但不能编辑文字。 |

### 快速开始

```bash
# 可编辑版本
python3 -m pip install --user python-pptx pillow playwright
python3 -m playwright install chromium
python3 scripts/html_to_pptx_editable.py deck.html deck.pptx

# 整图版本
brew install poppler
scripts/html_to_pptx_raster.sh deck.html deck.pptx
```

### 触发词（作为 Claude / Copilot skill 时会自动识别）

把 HTML 幻灯片转成 ppt · HTML 转 PPT · 网页转 ppt · 幻灯片下载成 PowerPoint · 下载成 ppt · 导出 ppt · 导出 PowerPoint · 生成可编辑的 ppt · 可编辑版本 · editable 版本 · 保留原样的 ppt · 整图版 ppt · 把这个 deck 转成 ppt · 这个网页幻灯片能导出成 ppt 吗

### 怎么选

- 后面还要改文案 / 换图 / 拆到别的 deck → **可编辑**
- 有渐变、渐变文字、SVG 图表、精细动画、只为分享/打印 → **整图**

### 中文字体注意

- **可编辑模式** 直接复用网页里的字体名。看的人电脑上要装同名 CJK 字体（`Noto Sans SC` / `PingFang SC` / `Source Han Sans` / `SimSun` / `SimHei`），否则 PowerPoint 会回退。
- **整图模式** 字体直接烧进像素里，对方无需安装字体。

---

## Quick start

### Editable mode

```bash
python3 -m pip install --user python-pptx pillow playwright
python3 -m playwright install chromium

python3 scripts/html_to_pptx_editable.py deck.html deck.pptx
```

### Raster mode

```bash
brew install poppler   # for pdftoppm
python3 -m pip install --user python-pptx

scripts/html_to_pptx_raster.sh deck.html deck.pptx
```

## When to use which

- **Editable** — you want to keep tweaking copy, restyle text, swap out images, or reuse slides in another deck.
- **Raster** — the deck depends on gradients, custom SVG charts, or animation timing that must render exactly as authored.

## How editable mode works

1. Launches headless Chromium at 1920 × 1080.
2. For every `.slide`, activates it (adds `.active .visible`, ends reveal animations), then walks the DOM.
3. Emits one native PowerPoint object per visual leaf:
   - Text-leaf blocks → editable text boxes with per-run font family, size, weight, italic, color, letter-spacing, alignment, line-height, and text-transform.
   - `<img>` → picture. WEBP / AVIF are auto-converted to PNG via Pillow. Extension-vs-content mismatches are sniffed.
   - Elements with a solid `background-color` or borders → rounded/square shape with matching fill + line.
   - Gradient backgrounds, gradient-clipped text, and `<svg>` roots → captured as a **discrete picture object** (still individually selectable, replaceable, movable — just not text-editable).
4. Assembles a 16:9 `.pptx` at `1920 × 9525 EMU × 1080 × 9525 EMU` on the blank layout.

## HTML deck conventions

The extractor expects:

- Each slide is an element with class `slide`, sized `1920 × 1080 px`, laid out as `position:absolute; inset:0` inside a `.deck-stage` wrapper.
- Slide visibility is toggled by adding/removing `.active` and `.visible`.
- Reveal animations use `.reveal` and `.d-stagger > *` starting at `opacity: 0`, promoted to visible when the parent `.slide` has `.visible`.

If your deck uses a different scheme, adjust the `ACTIVATE_JS` block in `scripts/html_to_pptx_editable.py` — that is the only slide-aware code.

## Non-Latin (CJK) text

- **Editable mode** copies `computed-style.fontFamily` straight through. The reader's PowerPoint needs the primary CJK font installed (`Noto Sans SC`, `PingFang SC`, `Source Han Sans`, `SimSun`, `SimHei`…), otherwise PowerPoint falls back to a system CJK font.
- **Raster mode** bakes fonts into the pixels — no client-side font install needed.

## Known caveats (editable mode)

- Gradient backgrounds are captured as a picture that includes any descendant text; the editable text is then overlaid on top in the correct position. It looks identical, but if you edit the text the old text remains visible inside the picture underneath. Delete the picture if you need to fully rewrite a gradient block.
- Inline styling (`<b>`, `<em>`, `<span style>`) is preserved as separate runs inside the same paragraph. Complex flex/grid alignment inside a single text-leaf can drift by 1–2 px.
- Occasional `element is not stable` screenshot timeouts on animated elements (about 1 per 50 slides). The build continues; that one decorative bitmap is skipped.

## Layout

```
html-to-ppt/
├── SKILL.md                          Claude / Copilot skill entry point
├── README.md
├── LICENSE
└── scripts/
    ├── html_to_pptx_editable.py      Playwright + python-pptx (native shapes)
    └── html_to_pptx_raster.sh        Chromium headless + pdftoppm + python-pptx
```

Both scripts are self-contained. Copy them into a project's `tools/` folder or run them straight from this repo — nothing here is project-specific.

## Using as a Claude / Copilot skill

Drop the folder into your skills directory:

```bash
git clone https://github.com/yanliudesign/html-to-ppt.git ~/.claude/skills/html-to-ppt
```

The skill will auto-trigger on phrases like *"download as PowerPoint"*, *"把 HTML 幻灯片转成 ppt"*, *"make an editable version of this deck"*, or *"export slides to PPTX"*.

## License

MIT
