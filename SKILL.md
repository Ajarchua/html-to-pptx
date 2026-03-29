---
name: html-to-pptx
version: 1.1.0
description: >
  Use this skill whenever a user wants to create a richly styled, fully editable PowerPoint presentation.
  This skill first generates a visually rich HTML preview, then converts it to a native PPTX file where
  all text, fonts, colors, shapes, tables, and charts are editable PowerPoint objects — not images.
  Trigger this skill when users say things like: "make me a presentation", "create a deck", "build slides",
  "I want a PPT with a nice design", "design a pitch deck", or any request where visual richness AND
  editability both matter. Prefer this skill over the standard pptx skill when the user cares about
  design quality. Also use this skill if the user explicitly mentions HTML-to-PPTX conversion.
---

# HTML → PPTX Skill

## 概述

本 skill 的核心思路：先用 HTML/CSS 生成视觉丰富的幻灯片预览，再将 HTML 的结构**纯元素映射**转换为原生 PPTX 对象。所有文字、形状、表格、图表均为可编辑的 PowerPoint 原生元素。

**完整流程：**
1. 理解用户需求，规划幻灯片内容与风格
2. 生成结构化 HTML（每张幻灯片一个 `<section>`）
3. 在 Artifact 中展示 HTML 预览，等待用户确认
4. 调用转换脚本，输出 `.pptx` 文件

---

## Step 0（可选）：提取用户模板风格

如果用户上传了 `.pptx` 模板文件，**在生成 HTML 之前**先运行风格提取：

```bash
python /home/claude/html-to-pptx/scripts/extract_style.py \
  --input /mnt/user-data/uploads/template.pptx \
  --output /home/claude/style_token.json
```

提取完成后，读取 `style_token.json`，并在生成 HTML 时严格遵循其中的参数：

| token 字段 | 用途 |
|---|---|
| `colors.background` | 幻灯片背景色 |
| `colors.primary` | 主色块、标题背景、侧栏颜色 |
| `colors.secondary` | 辅助色块 |
| `colors.accent` | 强调色、数据高亮 |
| `colors.text_primary` | 正文文字颜色 |
| `fonts.title.family/size_pt/bold` | 标题字体 |
| `fonts.body.family/size_pt` | 正文字体 |
| `layout.has_left_bar` | 是否保留左侧竖条装饰 |
| `layout.has_top_bar` | 是否保留顶部横条装饰 |
| `layout.left_bar_width_pct` | 左侧竖条宽度比例 |
| `decorations` | 主要装饰形状的位置/颜色规律 |

**关键原则：** token 提供的是品牌约束，不是模板复制。版式布局仍由 AI 根据内容自由创作，但配色、字体、装饰风格必须与 token 保持一致。

---

## Step 0：检查是否有用户上传的模板（可选）

如果用户上传了 `.pptx` 模板文件，**在生成 HTML 之前**先执行风格提取：

```bash
pip install python-pptx beautifulsoup4 lxml Pillow --break-system-packages -q

python /home/claude/html-to-pptx/scripts/extract_style.py \
  --input /mnt/user-data/uploads/template.pptx \
  --output /home/claude/style_profile.json
```

提取完成后读取 `style_profile.json`，将其中的配色、字体、装饰规律**完全替代**预设风格，用于后续 HTML 生成。没有上传模板则跳过此步骤，按 Step 1 自由创作。

---

## Step 1：生成 HTML

### 风格设计原则（无模板时）

**不要套用固定色板。** 每次根据内容主题从零设计，让配色、字体、版式为这个具体主题服务。

- **配色**：问自己"这个主题的颜色是什么"——科技感、自然感、严肃感、活力感各不同。选一个主导色占视觉面积60-70%，1-2个辅色，1个强调色点缀关键信息。深色背景适合封面和结尾，浅色适合内容页（"三明治"结构）。
- **字体**：标题要有个性，正文要清晰易读，两者形成对比。中文优先微软雅黑/思源黑体，英文优先Calibri/Georgia配对。
- **视觉动机**：选一个独特元素贯穿全程——圆角图片框、色圈图标、单侧粗色条——每张幻灯片都用它，形成统一感。
- **禁止**：标题下的装饰横线（AI感太强）、所有幻灯片用相同布局、纯文字+项目符号页。

### HTML 结构约定（严格遵守）

```html
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    /* 全局样式 */
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body { font-family: 'Microsoft YaHei', 'Calibri', sans-serif; }

    /* 每张幻灯片固定尺寸：1280×720px（16:9） */
    section.slide {
      width: 1280px;
      height: 720px;
      position: relative;
      overflow: hidden;
      /* 可设置背景色、渐变 */
    }

    /* 所有定位元素使用 position: absolute + px 坐标 */
  </style>
</head>
<body>
  <section class="slide" id="slide-1" data-layout="title">
    <!-- 幻灯片内容 -->
  </section>

  <section class="slide" id="slide-2" data-layout="content">
    <!-- 幻灯片内容 -->
  </section>
  <!-- 更多 slides... -->
</body>
</html>
```

### HTML 编写规则

**坐标系统（关键）：**
- 所有元素使用 `position: absolute` + `left/top/width/height`（px值）
- 幻灯片尺寸固定为 **1280×720px**
- 转换时坐标按比例映射到 PPTX（33.87cm × 19.05cm）
- 不要使用 flexbox/grid 做定位（可用于内部对齐，但外层容器必须绝对定位）

**支持的元素及其 data 属性：**

| 元素类型 | HTML写法 | 必要属性 |
|---|---|---|
| 标题/文字 | `<div class="text-box">` | `data-type="text"` |
| 图片 | `<img>` | `src`（URL或base64） |
| 矩形 | `<div class="shape">` | `data-type="rect"` |
| 圆形 | `<div class="shape">` | `data-type="ellipse"` |
| 箭头 | `<div class="shape">` | `data-type="arrow" data-direction="right/left/up/down"` |
| 表格 | `<table class="pptx-table">` | — |
| 图表 | `<div class="placeholder">` | `data-chart-type` + `data-chart` |

**文字元素示例：**
```html
<div class="text-box" data-type="text"
     style="position:absolute; left:80px; top:60px; width:800px; height:80px;
            font-size:48px; font-weight:bold; color:#1a1a2e; font-family:'Microsoft YaHei';">
  这是幻灯片标题
</div>
```

**形状元素示例：**
```html
<!-- 矩形色块 -->
<div class="shape" data-type="rect"
     style="position:absolute; left:0; top:0; width:400px; height:720px;
            background:#1a1a2e;">
</div>

<!-- 圆形装饰 -->
<div class="shape" data-type="ellipse"
     style="position:absolute; left:1100px; top:50px; width:120px; height:120px;
            background:rgba(255,200,0,0.3);">
</div>
```

---

### ⚠️ KPI 卡片必须用容器包裹写法

这是最容易出错的地方。**绝对不要**把背景色块和文字写成两个独立的绝对定位元素——坐标映射时会产生偏移，导致数字与卡片错位。

**❌ 错误写法（背景和文字分离）：**
```html
<!-- 背景色块 -->
<div class="shape" data-type="rect"
     style="position:absolute; left:40px; top:120px; width:260px; height:130px; background:#1e40af;">
</div>
<!-- 文字单独定位，坐标很容易和色块对不上 -->
<div class="text-box" data-type="text"
     style="position:absolute; left:60px; top:140px; width:220px; height:60px;
            font-size:48px; font-weight:bold; color:#ffffff;">
  $4.82B
</div>
```

**✅ 正确写法（容器包裹，背景+文字在同一元素内）：**
```html
<div class="shape" data-type="rect"
     style="position:absolute; left:40px; top:120px; width:260px; height:150px;
            background:#1e40af;">
  <div class="text-box" data-type="text"
       style="padding: 16px 20px 0 20px; font-size:13px; color:#93c5fd;">
    2024年总收入
  </div>
  <div class="text-box" data-type="text"
       style="padding: 6px 20px 0 20px;
              font-size:52px; font-weight:bold; color:#ffffff; line-height:1.1;">
    $4.82B
  </div>
  <div class="text-box" data-type="text"
       style="padding: 4px 20px 0 20px; font-size:12px; color:#bfdbfe;">
    ▲ 28.4% YoY
  </div>
</div>
```

**容器包裹的四条规则：**
1. 外层 `<div class="shape">` 负责背景色和绝对定位坐标
2. 内层文字用 `padding` 控制位置，**不使用 `position:absolute`**
3. 外层高度要留足余量，宁多勿少（比内容总高多 20px 以上）
4. 适用于所有"色块背景+文字叠加"的组合：KPI卡片、标签、说明框、角标

**图表元素（推荐：placeholder 机制）：**

用 `class="placeholder"` 标记图表占位区域，转换脚本用 python-pptx 图表 API 精确填入，坐标还原更准确：

```html
<div class="placeholder" data-chart-type="bar"
     style="position:absolute; left:640px; top:150px; width:560px; height:400px;
            background:#eeeeee;"
     data-chart='{
       "title": "季度销售额",
       "categories": ["Q1","Q2","Q3","Q4"],
       "series": [
         {"name":"2023","values":[1200,1450,1380,1620]},
         {"name":"2024","values":[1500,1780,1920,2100]}
       ],
       "show_legend": true
     }'>
</div>
```

`data-chart-type` 支持：`bar`（柱状）、`line`（折线）、`pie`（饼图）、`bar-horizontal`（条形）、`area`（面积）。

HTML 预览阶段可用灰色背景 + 文字描述占位，转换后会替换为真实图表：

```html
<!-- 预览占位内容（转换时被替换） -->
<div class="placeholder" data-chart-type="line" ...>
  <span style="color:#999; font-size:14px;">📊 折线图：月度趋势</span>
</div>
```

**表格示例：**
```html
<table class="pptx-table"
       style="position:absolute; left:80px; top:200px; width:1100px;">
  <thead>
    <tr style="background:#1a1a2e; color:white;">
      <th style="padding:12px; font-size:16px;">指标</th>
      <th style="padding:12px; font-size:16px;">数值</th>
    </tr>
  </thead>
  <tbody>
    <tr style="background:#f5f5f5;">
      <td style="padding:10px; font-size:14px;">收入</td>
      <td style="padding:10px; font-size:14px;">¥2.4M</td>
    </tr>
  </tbody>
</table>
```

### 设计风格指引

参考 [design-guide.md](design-guide.md) 获取配色、版式、字体建议。核心原则：
- 每张幻灯片必须有视觉焦点（大色块/装饰形状/图表）
- 标题字号 40-56px，正文 16-24px
- 使用 `data-layout` 标注版式类型：`title` / `content` / `two-column` / `data` / `closing`

---

## Step 2：展示 HTML 预览

生成 HTML 后，**必须在 Artifact 中渲染预览**，并告知用户。

- 若使用了用户模板风格，说明："已识别您的模板风格（主色 `#xxx`，字体 xxx），按此风格生成了预览。"
- 若自由创作，说明设计思路："配色选用了 xxx 风格，因为……"

> "确认内容和风格后，我将转换为可编辑的 PPTX 文件。如需调整请告诉我。"

等待用户明确确认（"好的" / "转换" / "可以"）后，再进入 Step 3。

---

## Step 3：转换为 PPTX

### 安装依赖

```bash
pip install python-pptx beautifulsoup4 lxml requests Pillow --break-system-packages -q
```

### 转换脚本

将 HTML 保存到文件，然后调用转换脚本：

```bash
# 保存HTML
cat > /home/claude/slides.html << 'EOF'
[HTML内容]
EOF

# 执行转换
python /home/claude/html-to-pptx/scripts/convert.py \
  --input /home/claude/slides.html \
  --output /home/claude/output.pptx

# 移动到输出目录
cp /home/claude/output.pptx /mnt/user-data/outputs/presentation.pptx
```

转换脚本详见 [scripts/convert.py](scripts/convert.py)。

---

## Step 4：QA 验证

转换完成后执行：

```bash
# 文字内容验证
python -m markitdown /home/claude/output.pptx

# 转为图片做视觉对比
python scripts/office/soffice.py --headless --convert-to pdf /home/claude/output.pptx
pdftoppm -jpeg -r 150 /home/claude/output.pdf /home/claude/slide
ls -1 "$PWD"/home/claude/slide-*.jpg
```

检查要点：
- 文字是否完整、无截断
- 颜色是否正确映射
- 表格结构是否完整
- 图表数据是否正确

---

## 降级处理规则

| HTML 特性 | PPTX 处理方式 |
|---|---|
| CSS 渐变背景 | 转为纯色（取渐变起始色） |
| CSS 动画/transition | 静默忽略，转为静态 |
| Web字体（Google Fonts等） | 降级为微软雅黑/Calibri |
| border-radius（圆角） | 矩形保留圆角（pptx支持），极端值取近似 |
| box-shadow | 忽略 |
| SVG 元素 | 转为图片插入（不可拆解编辑） |
| `<canvas>` | 忽略 |

---

## 参考文件

- [design-guide.md](design-guide.md) — 设计原则与版式参考
- [scripts/convert.py](scripts/convert.py) — HTML→PPTX 核心转换脚本
- [scripts/extract_style.py](scripts/extract_style.py) — 模板风格提取脚本（用户上传模板时调用）
