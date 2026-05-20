# 05 — 设计模式：Hero 页 / 三色配色 / 标题精修

从 BambuLab 媒介策略 / Moloco CTV 等项目复盘出的中文 PPT 提案视觉模式。

## Pattern 1：三色配色锁定

中文提案最常见的失败是配色花，看起来不专业。**只用三色**，建立明确的"主色 / 辅色 / 强调色"层级。

### 已验证的配色方案

#### 科技深空（适合品牌发布、媒介策略）
```
主色：深蓝黑  #062032
辅色：亮绿    #05B040
强调：橙色    #FBAE40
```
- 深蓝黑做封面、节标题页、Hero 大字页背景
- 亮绿做关键 KPI、按钮、CTA、强调文本
- 橙色用于警示信息、对比项

#### 极简灰绿（适合产品介绍、客户提案）
```
主色：深灰    #1F2937
辅色：薄荷绿  #10B981
强调：暖黄    #F59E0B
```

#### 商务深蓝（适合行业报告、年度规划）
```
主色：海军蓝  #1E40AF
辅色：天蓝    #60A5FA
强调：金色    #FBBF24
```

### 配色铁律

1. **一个色占 60–70% 视觉重量**（通常是主色），不能三色平均分
2. **背景与正文要有对比**：深色背景配浅色文字，浅色背景配深色文字
3. **强调色出现频率 ≤ 10%**，每张内容页最多 1–2 处
4. **不要 4 种颜色以上**，加再多看起来都廉价

### Python 中的配色定义

```python
from pptx.dml.color import RGBColor

class Palette:
    PRIMARY   = RGBColor.from_string("062032")  # 深蓝黑
    SECONDARY = RGBColor.from_string("05B040")  # 亮绿
    ACCENT    = RGBColor.from_string("FBAE40")  # 橙
    BG        = RGBColor.from_string("FFFFFF")
    TEXT      = RGBColor.from_string("333333")
    TEXT_MUTE = RGBColor.from_string("8A8A8A")
    GRAY_FILL = RGBColor.from_string("F5F5F5")
    GRAY_LINE = RGBColor.from_string("E0E0E0")
```

## Pattern 2：大字 Hero 页（区域策略 / 章节高潮）

在长 deck（>15 张）里，每隔几张内容页插入一张"Hero 大字页"，能极大提升节奏感和阅读体验。

### 何时用
- 三个区域策略（美国 / 欧洲 / 亚太）的引出页
- 重要节点（"上半年"→"下半年"）
- KPI 总目标的引出
- 章节高潮（最重要的那一句话）

### 视觉模板

```
┌──────────────────────────────────┐
│                                  │
│         美 国                    │  ← 96–120pt 超大字
│         ▔▔▔▔                    │
│                                  │
│   高净值订阅家庭 + 节日采购旺季  │  ← 28–36pt 副标题
│                                  │
│      重点投：CTV 黄金时段        │  ← 22pt 三行要点
│      节奏：10 月预热 / 11 月爆发 │
│      预算占比：45%               │
│                                  │
└──────────────────────────────────┘
```

### Python 实现骨架

```python
def slide_region_hero(prs, region_zh, region_en, key_insight, points):
    layout = prs.slide_layouts[7]  # 空白
    slide = prs.slides.add_slide(layout)
    
    # 深色背景
    add_full_bg(slide, Palette.PRIMARY)
    
    # 超大字区域名
    add_textbox(slide, x=Inches(0.8), y=Inches(0.8),
                w=Inches(8), h=Inches(2.5),
                text=region_zh, font="微软雅黑", size=110, bold=True,
                color="FFFFFF")
    
    # 横线（亮绿色 0.6 英寸宽）
    add_line(slide, x=Inches(0.85), y=Inches(3.4),
             w=Inches(0.6), color=Palette.SECONDARY, weight=4)
    
    # 副标题（核心洞察）
    add_textbox(slide, x=Inches(0.85), y=Inches(3.7),
                w=Inches(11), h=Inches(0.6),
                text=key_insight, size=32, bold=False,
                color="FFFFFF")
    
    # 三行要点
    for i, pt in enumerate(points):
        add_textbox(slide, x=Inches(0.85), y=Inches(4.6 + i * 0.6),
                    w=Inches(11), h=Inches(0.5),
                    text=pt, size=20, color="CCCCCC")
```

## Pattern 3：信息卡片（替代项目符号）

不要用 PowerPoint 默认的"圆点 + 文字"列表。用卡片：

```
┌─────────┐  ┌─────────┐  ┌─────────┐
│   ①     │  │   ②     │  │   ③     │
│ 标题1   │  │ 标题2   │  │ 标题3   │
│         │  │         │  │         │
│ 描述文字 │  │ 描述文字 │  │ 描述文字 │
└─────────┘  └─────────┘  └─────────┘
```

### Python 实现

```python
def add_card(slide, x, y, w, h, *,
             num=None, title="", body="",
             accent_color=Palette.SECONDARY):
    # 卡片背景（浅灰）
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    card.fill.solid()
    card.fill.fore_color.rgb = Palette.GRAY_FILL
    card.line.color.rgb = Palette.GRAY_LINE
    
    # 编号圆形
    if num:
        circle = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            x + Inches(0.3), y + Inches(0.3),
            Inches(0.5), Inches(0.5))
        circle.fill.solid()
        circle.fill.fore_color.rgb = accent_color
        set_text(circle, str(num), font="微软雅黑", size=16,
                 bold=True, color="FFFFFF", align="center")
    
    # 卡片标题
    title_box = slide.shapes.add_textbox(
        x + Inches(0.3), y + Inches(0.95),
        w - Inches(0.6), Inches(0.5))
    set_text(title_box, title, size=14, bold=True, color="333333")
    
    # 卡片正文
    body_box = slide.shapes.add_textbox(
        x + Inches(0.3), y + Inches(1.5),
        w - Inches(0.6), h - Inches(1.7))
    set_text(body_box, body, size=11, color="555555")
```

## Pattern 4：底部 Banner（每页固定信息）

在内容密集的提案里，每页底部加一条 Banner（带品牌色 + 一句概括），能强化记忆点：

```
┌──────────────────────────────────┐
│                                  │
│  （主内容区域）                  │
│                                  │
│                                  │
├──────────────────────────────────┤
│ ▌ 通过精准媒介组合，触达 5M+ 高净值用户 │  ← 绿色左边框 + 一句话
└──────────────────────────────────┘
```

```python
def add_bottom_banner(slide, text):
    # 整条横向底色块
    bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0), Inches(7.0),
        Inches(13.33), Inches(0.5))
    bar.fill.solid()
    bar.fill.fore_color.rgb = Palette.GRAY_FILL
    bar.line.fill.background()
    
    # 左侧 4mm 绿色竖条
    accent = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0), Inches(7.0),
        Inches(0.05), Inches(0.5))
    accent.fill.solid()
    accent.fill.fore_color.rgb = Palette.SECONDARY
    accent.line.fill.background()
    
    # 文字
    tb = slide.shapes.add_textbox(
        Inches(0.2), Inches(7.05),
        Inches(13), Inches(0.4))
    set_text(tb, text, size=11, color="333333")
```

## Pattern 5：对比表格（双方案 / 前后对比）

不要用 PowerPoint 内置的"光滑炫酷"表格样式。重新画一个简洁的：

```
┌────────────┬──────────┬──────────┐
│  维度      │  方案 A  │  方案 B  │
├────────────┼──────────┼──────────┤
│  覆盖人群  │  500万   │  800万 ✓ │
├────────────┼──────────┼──────────┤
│  CPM       │  $9      │  $12     │
├────────────┼──────────┼──────────┤
│  受众精准  │  ⭐⭐⭐  │  ⭐⭐⭐⭐⭐ │
└────────────┴──────────┴──────────┘
```

- 第一列：维度名（深色背景 + 白字）
- 内容列：白底 / 极浅灰交替（zebra striping）
- 强调单元格：浅绿色背景或对勾标记

详见 [`examples/minimal_deck.py`](../examples/minimal_deck.py) 中的对比表格示例。

## Pattern 6：避免"AI 生成感"的设计

某些视觉元素是 AI 生成 PPT 的标志，会让客户一眼看穿。**禁用清单**：

❌ 标题下方的横线 / 装饰短线（特别是绿色的"动感横线"）
❌ 每张图都有的渐变色块
❌ 圆角矩形 + 阴影的卡片（过于"模板化"）
❌ 底部那种"|"竖线分隔的导航栏（看起来像 PPT 模板自带）
❌ 标题与正文中间夹一句英文翻译

替代方案：
✅ 用粗体字号差建立层级（不靠装饰线）
✅ 用纯色背景而非渐变
✅ 卡片用浅灰底 + 极细 1px 描边
✅ 留白比加装饰元素更显高级

---

上一篇：[04-cjk-fonts.md](04-cjk-fonts.md) | 下一篇：[06-qa.md](06-qa.md)
