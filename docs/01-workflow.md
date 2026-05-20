# 01 — Word/DOCX/Markdown → PPT 完整工作流

把一份非 PPT 格式的内容（Word 提案、Markdown 章节、纯文本）映射到 PPTX，是这个 skill 最高频的入口场景。

## 场景

- 客户给了 `.docx` 提案，需要做成 PPT 提交
- 自己写了 Markdown 文档，需要排成幻灯片
- 现有英文 deck，要按中文版式重制

## 标准流程（5 步）

### 1. 内容提取

**优先方案：markitdown（如果已装）**

```bash
python3 -m markitdown 方案.docx > 方案.md
```

输出 Markdown，结构最完整。

**备用方案：直接解压 DOCX（零依赖）**

`.docx` 本质是 ZIP 包，可以绕过 `python-docx` 安装：

```bash
unzip -o 方案.docx -d ./docx_unpacked/
# 正文在 word/document.xml
```

然后用 `xml.etree` 或 `lxml` 解析：

```python
from lxml import etree
tree = etree.parse("docx_unpacked/word/document.xml")
ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
paragraphs = ["".join(t.text or "" for t in p.findall(".//w:t", ns))
              for p in tree.findall(".//w:p", ns)]
```

或直接用本 skill 的 [`scripts/extract_docx.py`](../scripts/extract_docx.py)。

**为什么不直接用 `python-docx`？** 在公司代理网络下经常装不上，而 `unzip` 一定有。

### 2. 内容结构化

把提取出的文本分块到 slide 级单位。常见模式：

| Word 段落特征 | 对应 PPT 元素 |
|---------------|---------------|
| 一级标题（Heading 1）| 章节分隔页（section divider）|
| 二级标题（Heading 2）| 内容页标题 |
| 小标题加粗 | 卡片标题或子节 |
| 列表项 | 项目符号或卡片 |
| 表格 | 真实 PPT 表格（**不要用 Rectangle+TextBox 伪表格**，详见 [08](08-shape-rebuild.md)）|
| 数据点（数字+单位）| 大字 KPI 块 |

**写一份 deck plan 再开工**。可以先输出一份 `deck_plan.md`：

```markdown
- Slide 01 封面：标题 / 副标题 / 客户 logo
- Slide 02 目录：4 章节
- Slide 03 节1分隔页：背景与挑战
- Slide 04 项目背景：3 段文字 + 数据点
- Slide 05 媒介挑战：3 列卡片
- ...
```

写好后再开始写 build 脚本，避免边写边乱。

### 3. 选模板 / 枚举 layouts

如果有参考 PPT 模板，先用 [`scripts/inspect_template.py`](../scripts/inspect_template.py) 枚举它的全部 slideLayout 与占位符索引：

```bash
python3 scripts/inspect_template.py 参考模板.pptx
```

输出示例：

```
Layout[0] "标题幻灯片"      placeholders: title(idx=0), body(idx=1)
Layout[1] "1_标题幻灯片"    placeholders: title(idx=0)
Layout[3] "节标题"          placeholders: title, body
Layout[4] "两栏内容"        placeholders: title, content×2
Layout[7] "空白"            完全自由
Layout[9] "图片与标题"      placeholders: title, picture
```

**关键经验**：
- `Layout[2]` 名字常叫"标题和内容"但**实际无占位符**，看到就跳过
- 各模板 layout 索引不一定一致，每次都要先枚举确认
- 内容页通常用 `Layout[1]`（仅 title）或 `Layout[7]`（空白），自由度高

### 4. 写 build 脚本

参照 [`examples/minimal_deck.py`](../examples/minimal_deck.py) 的骨架：

```python
from pptx import Presentation
prs = Presentation("template.pptx")

# 清空模板自带的示例 slides，保留 master & layouts
remove_all_slides(prs)

# 逐页构建
slide_cover(prs)
slide_section_divider(prs, "第一章 背景")
slide_content_page(prs, title="项目背景", body="...")
# ...

prs.save("output.pptx")
```

**核心 helper 函数**（每个项目都要有）：

| 函数 | 用途 |
|------|------|
| `set_text(shape, text, font, size, bold, color)` | 字体规范封装，详见 [04](04-cjk-fonts.md) |
| `add_real_table(slide, rows, cols, ...)` | 真实表格而非伪表格 |
| `set_title(slide, text)` | 写入 layout 的 title 占位符（保证位置一致）|
| `clear_non_title_shapes(slide)` | 清除非标题形状重绘内容（详见 [08](08-shape-rebuild.md)）|
| `insert_slide_at(prs, idx, layout)` | 任意位置插入 slide（详见 [03](03-python-pptx-gotchas.md)）|
| `normalize_punctuation(text)` | 中英文标点统一 |

### 5. 双轨 QA

**永远不要假设第一次构建就对**。详见 [06-qa.md](06-qa.md)。

```bash
# 结构 QA：边界 / 文本溢出 / 标题冲突
python3 scripts/qa_deep.py output.pptx

# 视觉 QA：PIL 渲染几何边界
python3 scripts/render_layout.py output.pptx qa_layout/

# 终极 QA：转 PDF/图片人眼看
soffice --headless --convert-to pdf output.pptx
pdftoppm -jpeg -r 150 output.pdf slide
```

## 反例

❌ **直接用 LibreOffice 打开 docx 复制粘贴到 PPT** — 字体丢失、版式错乱、表格变形
❌ **每页都新建一个 layout 0 的 slide 然后手动加 textbox** — 标题位置不一致，模板换皮就崩
❌ **把所有内容堆在一页"目录页"上** — 信息过载，客户看不见重点

## 速查：常见 deck 类型的 slide 数量

| 类型 | 张数范围 | 关键章节 |
|------|---------|---------|
| 媒介策略提案 | 18–24 | 背景 / 人群 / 媒介矩阵 / 区域策略 / KPI / 排期 |
| CTV 投放方案 | 12–16 | 平台优势 / 资源 / 定向 / 创意 / 预算 / 衡量 |
| 产品介绍 | 8–12 | 概述 / 核心能力 / 案例 / 价格 |
| 周报/月报 | 6–10 | 上期回顾 / 本期数据 / 下期计划 |

---

下一篇：[02-template-clone.md](02-template-clone.md) — 模板克隆构建法
