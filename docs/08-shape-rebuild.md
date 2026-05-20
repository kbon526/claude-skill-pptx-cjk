# 08 — 形状重建：清空非标题形状 / 真表格替代伪表格

修改既有 PPTX 的内容时，最稳的做法是**保留标题，清空其他形状，重新画**。本文讲清楚为什么以及怎么做。

## 反模式：Rectangle + TextBox 伪表格

很多客户给的模板里，"看起来像表格"的内容其实是 N 个 Rectangle（背景色块）+ N 个 TextBox（文字）叠出来的，不是真实的 PPT 表格对象。

### 症状

- 修改单元格内容时找不到表格对象，只能逐个 Rectangle/TextBox 修
- 列宽对不齐，因为每个 Rectangle 的 left 是手动摆的
- 文字与背景色块的 z-order 关系容易乱
- 用 `slide.shapes.tables` 遍历找不到（因为不是真表格）

### 案例（来自 V4 项目）

V4 PPTX 里所有看起来像"表格"的内容（人群矩阵、对比表、数据汇总），全都是 Rectangle + TextBox 浮层。结果：
- 内容稍长就溢出 Rectangle 边界
- 列对齐手抖一次就废
- 维护成本极高

### 解法：用真表格

```python
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

def add_real_table(slide, x, y, w, h, data,
                   *, header_bg="062032", header_fg="FFFFFF",
                   even_row_bg="F5F5F5", font="微软雅黑"):
    """
    创建真实 pptx.table 对象。
    data: 二维列表，第一行是 header
    """
    rows, cols = len(data), len(data[0])
    table_shape = slide.shapes.add_table(rows, cols, x, y, w, h)
    table = table_shape.table
    
    for r, row_data in enumerate(data):
        for c, cell_text in enumerate(row_data):
            cell = table.cell(r, c)
            cell.text = ""  # 清空默认
            tf = cell.text_frame
            p = tf.paragraphs[0]
            run = p.add_run()
            run.text = str(cell_text)
            run.font.name = font
            run.font.size = Pt(11)
            
            if r == 0:
                # Header 行
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor.from_string(header_bg)
                run.font.color.rgb = RGBColor.from_string(header_fg)
                run.font.bold = True
            elif r % 2 == 0:
                # 偶数行（zebra）
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor.from_string(even_row_bg)
                run.font.color.rgb = RGBColor.from_string("333333")
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor.from_string("FFFFFF")
                run.font.color.rgb = RGBColor.from_string("333333")
    
    return table_shape
```

### 真表格的限制

- 不支持单元格内多段落对齐独立设置（每个 cell 共享段落对齐）
- 不能合并跨多行 / 多列的视觉效果（合并是有的，但样式控制不灵活）
- cell 高度会随内容自动撑大，不能设 fixed height

如果要"自定义合并 + 跨列对齐"等高级效果，伪表格反而是合理选择，但要约定：**伪表格的所有元素 group 起来**（`slide.shapes.add_group_shape`），方便统一移动。

## 模式：clear_non_title_shapes 重建法

### 何时用

修改一张既有 slide 的内容时，比起"小心翼翼地修每个 textbox"，**清空重建**更可靠：
- 标题保留（位置、字体、占位符状态都对）
- 其他全部清掉
- 然后按 build 脚本重新画

### 实现

```python
def clear_non_title_shapes(slide):
    """删除所有非标题形状，保留标题占位符。"""
    shapes_to_remove = []
    for shape in slide.shapes:
        if shape.has_text_frame and shape.is_placeholder:
            ph_type = shape.placeholder_format.type
            if ph_type in (13, 14):  # TITLE, CENTER_TITLE
                continue  # 保留
        shapes_to_remove.append(shape)
    
    for shape in shapes_to_remove:
        sp = shape._element
        sp.getparent().remove(sp)
```

注：`13` 和 `14` 是 PP_PLACEHOLDER 的整数值。可以用枚举：

```python
from pptx.enum.text import PP_PLACEHOLDER
# PP_PLACEHOLDER.TITLE = 13
# PP_PLACEHOLDER.CENTER_TITLE = 14
```

### 完整重建 pattern

```python
def slide_us_strategy(prs, idx):
    """重建 idx 号 slide 的内容（保留标题）。"""
    slide = prs.slides[idx]
    clear_non_title_shapes(slide)
    
    # 重新画
    add_real_table(slide,
        Inches(0.5), Inches(1.5),
        Inches(12.3), Inches(5.0),
        data=[
            ["维度", "策略", "预算占比"],
            ["核心人群", "高收入家庭", "45%"],
            ["重点节点", "黑五感恩节", "30%"],
            ["创意重心", "CTV 黄金时段", "25%"],
        ])
    
    add_bottom_banner(slide, "美国市场聚焦节日采购旺季")
```

## 模式：group / ungroup

如果一组形状必须一起移动（如卡片的"背景 + 编号 + 标题 + 正文"），用 group：

```python
from pptx.shapes.graphfrm import GraphicFrame

def add_card_grouped(slide, x, y, w, h, *, num, title, body):
    shapes = []
    # 创建各组件
    bg = slide.shapes.add_shape(...)
    shapes.append(bg)
    circle = slide.shapes.add_shape(...)
    shapes.append(circle)
    title_box = slide.shapes.add_textbox(...)
    shapes.append(title_box)
    body_box = slide.shapes.add_textbox(...)
    shapes.append(body_box)
    
    # python-pptx 暂不支持 add_group_shape，需要手动操作 XML
    # 或者干脆不 group，用相对坐标管理
    return shapes
```

**注意**：python-pptx 1.0.x **不支持** `slide.shapes.add_group_shape()`。如果必须 group，要手动操作 XML：

```python
from pptx.oxml.ns import qn
from lxml import etree

def group_shapes(slide, shapes_to_group):
    sp_tree = slide.shapes._spTree
    grp = etree.SubElement(sp_tree, qn("p:grpSp"))
    # 设置 grp 的 nvGrpSpPr 和 grpSpPr ...
    for shape in shapes_to_group:
        sp_tree.remove(shape._element)
        grp.append(shape._element)
```

实际操作中**不推荐 group**，用相对坐标管理 + helper 函数封装更稳。

## 模式：z-order 控制

### 提到最前

```python
def bring_to_front(shape):
    sp_tree = shape._element.getparent()
    sp_tree.remove(shape._element)
    sp_tree.append(shape._element)
```

### 送到最后（背景）

```python
def send_to_back(shape):
    sp_tree = shape._element.getparent()
    sp_tree.remove(shape._element)
    # 找到第一个非 nvGrpSpPr/grpSpPr 的位置
    first_sp_idx = 0
    for i, child in enumerate(sp_tree):
        if not child.tag.endswith(("nvGrpSpPr", "grpSpPr")):
            first_sp_idx = i
            break
    sp_tree.insert(first_sp_idx, shape._element)
```

## 速查：何时用什么 pattern

| 场景 | 推荐做法 |
|------|----------|
| 表格数据 | `add_real_table`（真表格）|
| 自由布局多卡片 | 多个 `add_card` + 相对坐标 |
| 修改既有 slide 内容 | `clear_non_title_shapes` + 重建 |
| 数据可视化（柱图/线图）| `slide.shapes.add_chart` 真实图表 |
| 流程图 / 时间轴 | 多个 shape + connector，不要 group |
| 全屏背景 | `add_full_bg` 矩形 + `send_to_back` |

## 反例

❌ **用 Rectangle + TextBox 伪造表格**：维护噩梦，对齐易崩
❌ **修改 textbox 时不清空 text_frame，直接 append**：旧内容会留下来
❌ **依赖 PowerPoint 自动对齐**：python-pptx 生成的不会自动对齐，必须自己算坐标
❌ **同一组形状手动一个一个移**：用 helper 函数 + 基准坐标 + 偏移量

---

上一篇：[07-slidelayout-fix.md](07-slidelayout-fix.md) | 回到 [SKILL.md](../SKILL.md)
