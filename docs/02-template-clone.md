# 02 — 模板克隆构建法

不要从零创建 PPT。复用一个既有 PPTX 模板，是中文场景下保证视觉品质的最稳路径。

## 为什么"克隆模板"是最优解

从零创建（如 pptxgenjs）的问题：
- 中文字体很难配准（系统字体 vs 模板字体）
- 品牌色容易跑偏
- Master Slide 的装饰元素（logo、底纹、页码）必须手动重画
- 节标题页 / 封面 / 封底等异形版式从零搭非常累

克隆模板的好处：
- 视觉风格 100% 复用
- slideLayouts 已有占位符，写入即位置统一
- 主题字体（theme fonts）自动生效
- 装饰元素（logo、母版图形）零成本继承

## 完整流程

### 1. 解包模板看清楚结构

PPTX 是 ZIP 包，可以直接 unzip：

```bash
mkdir -p template_unpack
unzip -o template.pptx -d template_unpack/
```

关键文件：
| 路径 | 内容 |
|------|------|
| `ppt/theme/theme1.xml` | 主题字体（major/minor）、配色 |
| `ppt/slideMasters/slideMaster1.xml` | 母版（logo、底纹、页脚）|
| `ppt/slideLayouts/slideLayoutN.xml` | 各版式（占位符位置、默认字号）|
| `ppt/slides/slideN.xml` | 已有幻灯片（参考用）|

**先看 `theme1.xml` 的字体定义**：

```xml
<a:fontScheme name="自定义">
  <a:majorFont>
    <a:latin typeface="Calibri"/>
    <a:ea typeface="微软雅黑"/>  <!-- 标题字体 -->
  </a:majorFont>
  <a:minorFont>
    <a:latin typeface="Calibri Light"/>
    <a:ea typeface="微软雅黑"/>  <!-- 正文字体 -->
  </a:minorFont>
</a:fontScheme>
```

记住这两个字体名，后面写脚本时用。

### 2. 用 Python 枚举 slide_layouts

```python
from pptx import Presentation
prs = Presentation("template.pptx")
for i, layout in enumerate(prs.slide_layouts):
    placeholders = [(p.placeholder_format.idx, p.placeholder_format.type, p.name)
                    for p in layout.placeholders]
    print(f"Layout[{i}] {layout.name!r}: {placeholders}")
```

或直接用 [`scripts/inspect_template.py`](../scripts/inspect_template.py)。

**典型输出**（基于 BambuLab 模板的真实数据）：
```
Layout[0]  "标题幻灯片"   [(0, TITLE), (1, BODY)]
Layout[1]  "1_标题幻灯片"  [(0, TITLE)]
Layout[2]  "标题和内容"   []                       ← 空，慎用
Layout[3]  "节标题"       [(0, TITLE), (1, BODY)]
Layout[4]  "两栏内容"     [(0, TITLE), (1, OBJECT), (2, OBJECT)]
Layout[7]  "空白"        []                        ← 完全自由
Layout[9]  "图片与标题"   [(0, TITLE), (1, PICTURE)]
```

### 3. 清空模板示例 slides

模板里通常有几张示例幻灯片，要清掉但保留 master & layouts：

```python
def remove_all_slides(prs):
    """清空所有 slides，保留 layouts 和 master。"""
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst):
        sldIdLst.remove(sldId)
        rId = sldId.get(qn("r:id"))
        prs.part.drop_rel(rId)
```

或者更简单的版本（不释放 part 引用，文件略大但够用）：

```python
def remove_all_slides(prs):
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    for slide in slides:
        xml_slides.remove(slide)
```

### 4. 用 layout 创建新 slide + 写占位符

```python
def slide_section_divider(prs, title_text):
    layout = prs.slide_layouts[3]  # "节标题"
    slide = prs.slides.add_slide(layout)
    # 找到 title 占位符
    for ph in slide.placeholders:
        if ph.placeholder_format.type == 13:  # TITLE
            ph.text = title_text
            break
    return slide
```

**重点**：标题始终通过占位符写入，**不用自由 textbox**。这样：
- 位置由 layout 决定，全 deck 一致
- 字体自动继承主题
- 母版的装饰线、页码自动出现

### 5. 内容页用空白 layout + 自定义形状

对于内容复杂的页（数据卡片、流程图、表格），用 `Layout[7]`（空白）然后自由排版：

```python
def slide_content(prs, title, cards):
    layout = prs.slide_layouts[1]  # 标题占位符 + 自由布局
    slide = prs.slides.add_slide(layout)
    set_title(slide, title)
    # 自由摆放卡片
    for i, card in enumerate(cards):
        x = Inches(0.5 + i * 4.0)
        add_card(slide, x, Inches(2), Inches(3.5), Inches(3),
                 title=card["title"], body=card["body"])
    return slide
```

## 经常踩的坑

### 占位符 idx 不连续

不要假设 idx 是 0,1,2,3...，可能跳号。**总是按 type 而非 idx 找占位符**：

```python
# ❌ 错的
slide.placeholders[0]  # idx=0 不一定是 title

# ✅ 对的
for ph in slide.placeholders:
    if ph.placeholder_format.type == 13:  # TITLE
        ...
```

或者用 `from pptx.enum.text import PP_PLACEHOLDER` 的常量做判断。

### 标题占位符高度不够装中文长标题

模板的 `slideLayout.xml` 里 title 的 `<p:sp>/<p:spPr>/<a:xfrm>/<a:ext cy="..."/>` 可能太矮（<1.5cm），中文长标题塞不下。详见 [07-slidelayout-fix.md](07-slidelayout-fix.md)。

### 模板的 master 元素挡住内容

母版上有时有大块装饰图形（如全屏底色），新加的 textbox 可能被挡。解决方案：
- 把 textbox 的 z-order 提到最前
- 或选用没有大装饰元素的 layout

### 直接复制 slide 的 XML 不能跨 PPTX

PPTX 内部的 rId 关系（图片、图表的引用）会断。如果要跨文件复制 slide，用 [`copy-slide-py`](https://github.com/scanny/python-pptx/issues/132) 这类社区方案，或干脆把内容重新构建一遍。

## 模板从哪儿来

- 客户提供的既有 deck（最常见）
- 公司品牌库 / VI 系统下发的模板
- 用 PowerPoint 现做一个"骨架模板"（封面 + 节标题 + 内容页 + 封底各一张）
- 从开源资源找（[Slidesgo](https://slidesgo.com/) 等），但需谨慎检查中文字体是否友好

---

上一篇：[01-workflow.md](01-workflow.md) | 下一篇：[03-python-pptx-gotchas.md](03-python-pptx-gotchas.md)
