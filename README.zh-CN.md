[English](README.md) · **中文**

# html-to-ppt

把自包含的 HTML 幻灯片（1920 × 1080，每页是一个 `.slide` 元素）转成 `.pptx`。

## 为什么要把 HTML 转成 PowerPoint？

HTML deck 好看，但一旦你要**真正把它推出去用**，PowerPoint 在几乎每个实际场景里都更胜一筹：

- **更容易修改。** 大多数同事不会 CSS，但都会 PowerPoint。给他们一份 `.pptx`，PM、销售、老板都能 10 秒钟改个错别字、换个标题、替张截图，不用去翻 `<div class="…">`。
- **天生为团队协作而生。** PowerPoint 有修订记录、评论、多人协同、版本历史、幻灯片对比。HTML 文件丢到群里，每个人本地改一份，最后没法合并。
- **分发零阻力。** `.pptx` 是商务演示的通用货币 —— 邮件附件、SharePoint、Google Slides 导入、导出 PDF、打印讲义、嵌 Confluence / 飞书 / Notion 全都可以。HTML deck 要托管、要一个不会失效的链接、还要一个能把 CSS 渲染成你想要样子的浏览器。
- **设计资产可以复用。** 每一段文字、每种颜色、每张图片都是原生对象 —— 下一份 deck 里能复用，客户模板里能重混，换套品牌不用重做。
- **演讲工具开箱即用。** 演讲者备注、演讲者视图、激光笔、Teams / Zoom 无缝共享、Wi-Fi 差的会议室里能离线播放。
- **归档更放心。** `.pptx` 是开放的、有版本的、ISO 标准格式。HTML 依赖字体、CDN、浏览器差异、JS 库，全都会退化。2010 年的 PowerPoint 现在照样能打开；2010 年的网页大概率打不开了。

目标不是替换你的 HTML 源文件 —— 那份仍然是**母版**。目标是让这份 deck 在真正会被人编辑、转发、打印的工具里**多活一次**。

## 两种模式

| 模式 | 效果 |
|---|---|
| **Editable 可编辑** | 每段文字、每张图、每个色块都变成 **原生 PowerPoint 对象**，可以选中、改字、换图、复用到别的 deck。 |
| **Raster 整图** | 每页导出为一张整图，视觉 100% 保真，但不能编辑文字。 |

面向使用 `1920×1080` 固定舞台的 HTML deck —— 和 [frontend-slides](../frontend-slides) skill、大多数单文件 HTML 演示的写法一致。

## 快速开始

### 可编辑模式

```bash
python3 -m pip install --user python-pptx pillow playwright
python3 -m playwright install chromium

python3 scripts/html_to_pptx_editable.py deck.html deck.pptx
```

### 整图模式

```bash
brew install poppler   # 提供 pdftoppm
python3 -m pip install --user python-pptx

scripts/html_to_pptx_raster.sh deck.html deck.pptx
```

## 怎么选

- **可编辑** —— 后面还要改文案、换图、拆到别的 deck。
- **整图** —— 依赖渐变 / 渐变文字 / 自定义 SVG 图表 / 精细动画时序，只想快速分享或打印。

## 可编辑模式的工作原理

1. 用无头 Chromium 打开 HTML，视口 1920 × 1080。
2. 对每一个 `.slide`：先激活它（加 `.active .visible`、结束 reveal 动画），再遍历整个 DOM。
3. 每一个视觉叶子节点生成一个原生 PowerPoint 对象：
   - 纯文本块 → 可编辑文本框，保留字体、字号、字重、斜体、颜色、字距、对齐、行高、text-transform。
   - `<img>` → 图片。WEBP / AVIF 通过 Pillow 自动转 PNG；扩展名和真实格式不一致时会嗅探修正。
   - 有纯色背景或边框的元素 → 圆角/直角形状，保留填充+描边。
   - 渐变背景、渐变裁切文字、`<svg>` 根节点 → 作为**独立图片对象**截图（仍可单独选中、替换、移动，只是文字不可编辑）。
4. 用空白版式组装 16:9 `.pptx`：`1920 × 9525 EMU × 1080 × 9525 EMU`。

## HTML deck 结构约定

提取器假设：

- 每一页是一个 `class="slide"` 的元素，尺寸 `1920 × 1080 px`，`position:absolute; inset:0`，包在 `.deck-stage` 里。
- 通过增删 `.active` / `.visible` 切换页面。
- reveal 动画用 `.reveal` 和 `.d-stagger > *`，初始 `opacity: 0`，父级 `.slide` 拿到 `.visible` 后可见。

如果你的 deck 用别的方案，改 `scripts/html_to_pptx_editable.py` 里的 `ACTIVATE_JS` 块即可 —— 只有那里跟具体页面结构相关。

## 中文字体注意

- **可编辑模式** 直接原样复制网页里的 `font-family`。读者的 PowerPoint 需装同名 CJK 字体（`Noto Sans SC` / `PingFang SC` / `Source Han Sans` / `SimSun` / `SimHei`），否则会自动回退到系统字体。
- **整图模式** 字体直接烧进像素，读者无需安装字体。

## 已知限制（可编辑模式）

- 渐变背景截图时会把里面的文字也拍进去，然后再把可编辑文字覆盖在上面。视觉一致，但你如果编辑了上面那层文字，底图里的旧文字仍在。要整块改文字请先删掉那张底图。
- 内联样式（`<b>` / `<em>` / `<span style>`）会保留为同一段落里的多个 run。复杂的 flex/grid 对齐在单个文本叶子里可能有 1–2 px 漂移。
- 偶尔（约 1/50 页）动画中的元素会触发 `element is not stable` 截图超时，构建继续，该装饰位图跳过。

## 目录结构

```
html-to-ppt/
├── SKILL.md                          Claude / Copilot skill 入口
├── README.md                         英文版
├── README.zh-CN.md                   中文版
├── LICENSE
└── scripts/
    ├── html_to_pptx_editable.py      Playwright + python-pptx（原生形状）
    └── html_to_pptx_raster.sh        Chromium headless + pdftoppm + python-pptx
```

两个脚本都自包含，可以直接复制到项目的 `tools/` 目录，或者从这个仓库里直接跑 —— 里面没有任何项目特定的东西。

## 作为 Claude / Copilot skill 使用

克隆到你的 skills 目录：

```bash
git clone https://github.com/yanliudesign/html-to-ppt.git ~/.claude/skills/html-to-ppt
```

skill 会自动在以下触发词上激活：

> 把 HTML 幻灯片转成 ppt · HTML 转 PPT · 网页转 ppt · 幻灯片下载成 PowerPoint · 下载成 ppt · 导出 ppt · 导出 PowerPoint · 生成可编辑的 ppt · 可编辑版本 · editable 版本 · 保留原样的 ppt · 整图版 ppt · 把这个 deck 转成 ppt · 这个网页幻灯片能导出成 ppt 吗
>
> download as PowerPoint · make an editable version of this deck · export slides to PPTX

## License

MIT
