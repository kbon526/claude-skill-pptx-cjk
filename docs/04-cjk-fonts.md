# 04 — 中文字体规范与 CJK 排版

中文 PPT 的字体问题是最高频翻车点。这篇定下铁律。

## 字体选择三铁律

### 律 1：标题用微软雅黑 Bold，正文用微软雅黑 Light

不要用 **Regular**，原因：
- 在 PPT 里 Regular 字重在标题位置看起来不够分量
- 在正文位置又比 Light 更"糊"（视觉重量过重）
- Bold 与 Light 的对比清晰，建立明确的视觉层级

| 元素 | 字体 | 字号 | 字重 |
|------|------|------|------|
| 封面主标题 | 微软雅黑 Bold | 40–54pt | Bold |
| 封面副标题 | 微软雅黑 Light | 18–22pt | Light |
| 节标题页 | 微软雅黑 Bold | 36–48pt | Bold |
| 内容页大标题 | 微软雅黑 Bold | 24–28pt | Bold |
| 卡片小标题 | 微软雅黑 Bold | 14–16pt | Bold |
| 正文 | 微软雅黑 Light | 11–13pt | Light |
| 注脚 / 引用 | 微软雅黑 Light | 8–10pt | Light |
| 大数字 KPI | 微软雅黑 Bold | 48–72pt | Bold |
| 大数字单位 | 微软雅黑 Light | 14–18pt | Light |

### 律 2：英文与中文用同一族字体（避免混排断层）

中英混排时不要让英文跑去 Calibri 而中文留在微软雅黑。两个解决方案：

**方案 A：英文也用微软雅黑**
- 简单直接
- 缺点：英文字符在微软雅黑下没那么好看

**方案 B：中文用微软雅黑，英文用 Calibri，但在 run.font 上**同时**设置 `name`（西文）和 `name_east_asian`（CJK）

```python
run.font.name = "Calibri"  # 西文
# python-pptx 1.0+ 不直接支持 east_asian，需要操作 XML
rPr = run._r.get_or_add_rPr()
ea = OxmlElement("a:ea")
ea.set("typeface", "微软雅黑")
rPr.append(ea)
```

**推荐方案 A**，简单不出错。

### 律 3：禁用 PingFang / 苹方 / 华文系列

虽然 PingFang 在 macOS 上很美，但在 Windows 上没有，会回退到默认中文字体（宋体），导致客户在 Windows 上看到的版式与你设计的完全不一样。

**只用微软雅黑**（Windows / macOS / 在线 Office 都有）。

如果要更正式的氛围：
- 可以考虑 **思源黑体（Source Han Sans / Noto Sans CJK SC）**，开源跨平台
- 但要确保客户机器装了，否则也会回退

## CJK 字符宽度估算（用于文本溢出预警）

中文字符在排版上是"全宽"的，宽度估算与拉丁字符不同：

```python
def estimate_text_width_inches(text, font_pt):
    """返回文字总宽度（英寸）。"""
    cjk_count = sum(1 for c in text if '一' <= c <= '鿿'
                                       or '　' <= c <= '〿'
                                       or '＀' <= c <= '￯')
    latin_count = len(text) - cjk_count
    # CJK 全宽 = font_pt / 72 英寸
    # Latin 半宽 ≈ font_pt × 0.55 / 72 英寸
    return (cjk_count * font_pt + latin_count * font_pt * 0.55) / 72
```

`qa_deep.py` 用这个公式估算文本是否会溢出 textbox。详见 [06-qa.md](06-qa.md)。

## set_text 封装（统一字体规范的关键）

```python
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

FONT_TITLE = "微软雅黑"
FONT_BODY = "微软雅黑"

def set_text(shape, text, *, font=FONT_BODY, size=14,
             bold=False, color="333333", align="left",
             line_spacing=1.2):
    tf = shape.text_frame
    tf.word_wrap = True
    tf.clear()
    p = tf.paragraphs[0]
    p.alignment = {
        "left": PP_ALIGN.LEFT,
        "center": PP_ALIGN.CENTER,
        "right": PP_ALIGN.RIGHT,
    }[align]
    p.line_spacing = line_spacing
    run = p.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor.from_string(color)
```

完整版见 [`scripts/set_text.py`](../scripts/set_text.py)，含多段落、混合字重等高级用法。

## 中文标题精修原则

中文标题不像英文那样可以靠"动词+宾语"短语就成立，**容易冗长**。下面是从 V5 项目里复盘出的精修规则：

### 规则 1：去除冗余英文小标题

如果标题下方有英文翻译副标题（如"项目背景 / Project Background"），通常可以删掉。中文受众不需要英文重复，留着只会挤占视觉空间。

### 规则 2：删除口语词

| 冗余 | 精修后 |
|------|--------|
| 高 ROI 的人群画像 | 人群画像 |
| 全面的市场背景分析 | 市场背景 |
| 我们如何思考媒介选择 | 媒介选择思路 |

### 规则 3：标题字数控制

- 内容页标题 ≤ 14 个字（含标点）
- 节标题页 ≤ 10 个字
- 封面主标题 ≤ 16 个字

超长就考虑拆成"主标题 + 副标题"两行。

### 规则 4：避免"的""了""和"等虚词

| 啰嗦 | 精修 |
|------|------|
| 中国市场的人群洞察和趋势分析 | 中国市场人群洞察 |
| 关于 CTV 媒介投放的预算分配方案 | CTV 预算分配方案 |

## 标点统一

中英文混排时标点容易混乱。建议在 build 脚本里加一个统一函数：

```python
def normalize_punctuation(text):
    """中文文本中的英文标点替换为中文标点。"""
    # 注意：仅在前后字符为中文时才替换
    ...
    return text
```

或者反过来，全 deck 强制用英文标点，保证统一。

## Hero 大字页的字号选择

V5 项目里实践过的"区域策略 Hero 页"配置：

| 元素 | 字号 |
|------|------|
| Hero 大标题（如"美国"）| 96–120pt |
| Hero 副标题 | 28–36pt |
| 三个核心要点 | 18–22pt |

字号差距要拉开（最大 / 最小 ≥ 4 倍），层级感才出得来。

详见 [05-design-patterns.md](05-design-patterns.md) 的 Hero 页 pattern。

---

上一篇：[03-python-pptx-gotchas.md](03-python-pptx-gotchas.md) | 下一篇：[05-design-patterns.md](05-design-patterns.md)
