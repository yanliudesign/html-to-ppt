**English** · [中文](README.zh-CN.md)

# html-to-ppt

Convert self-contained HTML slide decks (1920 × 1080 slides authored as `.slide` elements) to `.pptx`.

## Why convert HTML → PowerPoint?

HTML decks look great, but the moment you want to *ship* them, PowerPoint wins on almost every practical axis:

- **Easier to edit.** Most teammates don't know CSS. They know PowerPoint. Handing over a `.pptx` means anyone — PM, sales, exec — can fix a typo, retitle a slide, or swap a screenshot in 10 seconds instead of hunting through `<div class="…">`.
- **Built for team collaboration.** PowerPoint has track-changes, comments, co-authoring, revision history, and a "compare" mode. Sharing an HTML file over Slack means everyone edits a different local copy and the diffs are unmergeable.
- **Frictionless distribution.** `.pptx` is the universal currency of business decks — email attachments, SharePoint, Google Slides import, PDF export, printable handouts, embedded in Confluence/Notion. HTML decks need hosting, a link that doesn't rot, and a viewer that renders your CSS the same way you did.
- **Design lives on.** Every text style, color, and image becomes a native object — reusable in the next deck, remixable in a client template, restyleable to a new brand without a rebuild.
- **Presenter tools work out of the box.** Speaker notes, presenter view, laser pointer, seamless Teams/Zoom sharing, offline playback in a boardroom with flaky Wi-Fi.
- **Long-term archive.** `.pptx` is an open, versioned, ISO-standard format. HTML depends on fonts, CDNs, browser quirks, and JS libraries that all decay. A PowerPoint from 2010 still opens; a webpage from 2010 usually doesn't.

The goal isn't to replace your HTML source — that stays the master. The goal is to give the deck a **second life** in the tool where it will actually be edited, forwarded, and printed.

## Two modes

| Mode | Result |
|---|---|
| **Editable** | Every text run, image, and colored box becomes a **native PowerPoint object** you can select, restyle, and edit. |
| **Raster** | One full-bleed image per slide. Fast, always visually faithful, but not editable. |

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
├── README.zh-CN.md                   中文版
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
