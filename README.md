# html-to-pptx

> An Agent Skill that generates richly styled, **fully editable** PowerPoint presentations — by designing slides in HTML first, then converting every element into native PPTX objects.

## Why HTML → PPTX?

Most AI-generated PowerPoints are either visually boring or locked into rigid templates. This skill takes a different approach:

1. **Design in HTML/CSS** — full creative freedom, any layout, any color scheme
2. **Convert to native PPTX** — every text box, shape, table, and chart becomes a real PowerPoint object you can click and edit

No screenshots. No images of slides. Everything stays editable.

---

## Compatible Agents

| Agent | Install location |
|---|---|
| Claude Code | `~/.claude/skills/` or `.claude/skills/` |
| OpenAI Codex CLI | `~/.codex/skills/` |
| Cursor | `.cursor/skills/` |
| Any agent supporting the [Agent Skills](https://agentskills.io) standard | Per agent docs |

---

## Install

**Via npx (recommended):**
```bash
npx skills add Ajarchua/html-to-pptx
```

**Manual:**
```bash
# Claude Code (global)
git clone https://github.com/Ajarchua/html-to-pptx ~/.claude/skills/html-to-pptx

# Claude Code (project)
git clone https://github.com/Ajarchua/html-to-pptx .claude/skills/html-to-pptx
```

---

## Usage

Once installed, just describe what you want:

```
Make me a 5-slide pitch deck for a SaaS product focused on supply chain automation.
```

```
Create a quarterly business review presentation with charts and KPI data.
```

```
Build a presentation using my company template (attach your .pptx file).
```

The agent will:
1. Design slides as HTML with a style tailored to your topic
2. Show you a preview for confirmation
3. Convert to a fully editable `.pptx` file

---

## Features

- **🎨 Topic-driven design** — colors, fonts, and layouts chosen for your specific content, not from a generic template
- **✏️ Fully editable output** — text, shapes, tables, and charts are all native PowerPoint objects
- **📊 Chart support** — bar, line, pie, area, and horizontal bar charts with real data
- **📋 Table support** — styled tables with header rows and alternating colors
- **🖼️ Image support** — local files or URLs
- **📐 Overflow detection** — warns if any element exceeds slide boundaries before converting
- **🎭 Template matching** — upload your own `.pptx` template and the skill extracts its color scheme and fonts to match your brand

---

## Supported Elements

| Element | HTML Convention |
|---|---|
| Text box | `<div class="text-box" data-type="text">` |
| Rectangle | `<div class="shape" data-type="rect">` |
| Ellipse | `<div class="shape" data-type="ellipse">` |
| Arrow | `<div class="shape" data-type="arrow">` |
| Image | `<img src="...">` |
| Table | `<table class="pptx-table">` |
| Chart | `<div class="placeholder" data-chart-type="bar" data-chart='{...}'>` |

All positioned elements use `position: absolute` with `px` coordinates on a **1280 × 720px** canvas.

---

## Template Matching

Upload a `.pptx` file to have the skill extract and apply its visual style:

```
Here's our company template (attach file). Create a 6-slide product update presentation.
```

The skill reads the first slide of your template to extract:
- Background and primary colors
- Accent colors and text colors
- Title and body font names and sizes
- Layout patterns (left sidebar, top bar, decorative shapes)

Your brand colors and fonts are applied to freshly designed slides — preserving the brand feel without copying the layout mechanically.

---

## HTML Structure Reference

Each slide is a `<section class="slide">` inside a single HTML file:

```html
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    section.slide { width: 1280px; height: 720px; position: relative; overflow: hidden; }
  </style>
</head>
<body>

  <section class="slide" id="slide-1" data-layout="title"
           style="background-color:#1a1a2e;">

    <!-- Color block -->
    <div class="shape" data-type="rect"
         style="position:absolute; left:0; top:0; width:400px; height:720px;
                background:#e94560;"></div>

    <!-- Title -->
    <div class="text-box" data-type="text"
         style="position:absolute; left:460px; top:200px; width:720px; height:100px;
                font-size:52px; font-weight:bold; color:#ffffff; font-family:'Microsoft YaHei';">
      Presentation Title
    </div>

    <!-- Chart placeholder -->
    <div class="placeholder" data-chart-type="bar"
         style="position:absolute; left:460px; top:320px; width:720px; height:360px;"
         data-chart='{
           "title": "Quarterly Revenue",
           "categories": ["Q1","Q2","Q3","Q4"],
           "series": [{"name":"2024","values":[1500,1780,1920,2100]}]
         }'>
    </div>

  </section>

</body>
</html>
```

---

## File Structure

```
html-to-pptx/
├── SKILL.md                  # Agent instructions (loaded on trigger)
├── README.md                 # This file
├── design-guide.md           # Design principles and layout reference
└── scripts/
    ├── convert.py            # HTML → PPTX converter
    └── extract_style.py      # Template style extractor
```

---

## Requirements

- Python 3.8+
- Packages installed automatically on first run:
  ```
  pip install python-pptx beautifulsoup4 lxml requests Pillow
  ```

---

## Changelog

### v1.0.0
- Initial release
- Full element mapping: text, shapes, images, tables, charts
- Placeholder-based chart mechanism for accurate positioning
- Overflow detection with per-element warnings
- Template style extraction from uploaded `.pptx` files
- Topic-driven design system (no fixed color palettes)
- Compatible with Claude Code, Codex CLI, Cursor, and all Agent Skills-compatible tools

---

## License

MIT License — free to use, modify, and distribute.
