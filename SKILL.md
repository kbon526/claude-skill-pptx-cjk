---
name: pptx-cjk
description: 中文 PPT 提案专精 skill。基于 python-pptx + 模板克隆路线，处理中文场景下的 PPTX 创建、修改与 QA。当用户要求"做一份提案 PPT/媒介策略 PPT/把 Word 转成 PPT/按某个模板重排 PPT/修改 PPTX 文件/中文幻灯片字体/微软雅黑/Hero 大字页/三色配色"等中文 PPT 任务时使用。也覆盖 python-pptx 操作、slideLayout XML 修复、占位符问题、PPTX 转图片做视觉 QA、CJK 字符宽度估算等场景。与官方 pptx skill 互补：官方走 pptxgenjs(JS) + markitdown 路线偏英文场景，本 skill 走 python-pptx + 模板克隆路线偏中文场景。
license: MIT
---

# pptx-cjk — 中文 PPT 提案专精 Skill

中文场景下的 PPTX 自动化制作与 QA。基于 python-pptx，沉淀自媒介策略、CTV 广告投放、客户提案等多个真实项目。

## 何时使用本 Skill

✅ 用 Word 文档 / Markdown 内容生成中文 PPT 提案
✅ 按某个既有 PPTX 的版式克隆出新 PPT（保留品牌风格）
✅ 修改既有 PPTX（删页/插页/改标题/换字体/重排形状）
✅ 中文字体出现"显示成宋体""字号不对"等问题需要修复
✅ 客户交付前需要做 PPTX 的结构与视觉双轨 QA

❌ 英文场景从零创建 → 优先用官方 `pptx` skill
❌ 仅需要把 PPTX 转成 PDF → `soffice --headless --convert-to pdf` 即可，不必加载本 skill

## Quick Reference

| 你想做什么 | 看哪份文档 / 用哪个脚本 |
|------------|-------------------------|
| 把 Word/Markdown 内容做成 PPT | [docs/01-workflow.md](docs/01-workflow.md) |
| 按某个模板克隆新 PPT | [docs/02-template-clone.md](docs/02-template-clone.md) |
| python-pptx 报错 / 环境问题 | [docs/03-python-pptx-gotchas.md](docs/03-python-pptx-gotchas.md) |
| 中文字体规范、字号、CJK 宽度 | [docs/04-cjk-fonts.md](docs/04-cjk-fonts.md) |
| 设计 Hero 页 / 配色 / 标题精修 | [docs/05-design-patterns.md](docs/05-design-patterns.md) |
| 提交前的双轨 QA | [docs/06-qa.md](docs/06-qa.md) |
| 标题被截断 / 字号继承不稳 | [docs/07-slidelayout-fix.md](docs/07-slidelayout-fix.md) |
| 重建形状 / 真表格替换伪表格 | [docs/08-shape-rebuild.md](docs/08-shape-rebuild.md) |
| 工具脚本 | [scripts/](scripts/) |
| 5 页最小示例 | [examples/minimal_deck.py](examples/minimal_deck.py) |

## 标准工作流（推荐路径）

```
[输入素材]                    [模板]
  Word/Markdown               已有 .pptx
       ↓                          ↓
   提取文本 -----+------+---- 解包+枚举 layouts
                 ↓      ↓
              [内容映射]
                 ↓
           克隆模板 → 写入占位符 → 重建形状
                 ↓
           [双轨 QA]
              ↓     ↓
        qa_deep   PIL render
                 ↓
              [交付 .pptx]
```

具体每一步的命令、脚本、参数详见各 docs 文档。

## 三条铁律（违反即翻车）

1. **macOS 用户必须显式使用 Python 3.13** 路径，系统默认的 `python3` 大概率不含 `python-pptx`
   ```bash
   /Library/Frameworks/Python.framework/Versions/3.13/bin/python3 build.py
   ```
   详见 [03-python-pptx-gotchas.md#python3-路径](docs/03-python-pptx-gotchas.md)

2. **脚本文件名严禁与 Python 标准库冲突**（特别是 `inspect.py` / `types.py` / `io.py`），会触发 lxml 循环导入

3. **中文字体严格使用微软雅黑 Bold（标题）+ Light（正文），禁用 Regular**
   通过 `set_text()` 封装统一注入，详见 [04-cjk-fonts.md](docs/04-cjk-fonts.md)

## 不同任务的入口建议

### Case A：从零做提案（有 Word 源 + 参考模板）

1. Read [01-workflow.md](docs/01-workflow.md) 全文
2. 用 `scripts/extract_docx.py` 提取 Word 内容
3. 用 `scripts/inspect_template.py` 枚举模板的 slideLayout 索引
4. 参照 [examples/minimal_deck.py](examples/minimal_deck.py) 写自己的 build_xxx.py
5. QA：跑 `scripts/qa_deep.py` 和 `scripts/render_layout.py`

### Case B：修改既有 PPT

1. Read [03-python-pptx-gotchas.md](docs/03-python-pptx-gotchas.md) 避坑
2. Read [08-shape-rebuild.md](docs/08-shape-rebuild.md) 理解形状操作模式
3. 如果要在中间插入 slide，用 `scripts/insert_slide_at.py`
4. 如果是标题被截断/字号问题，看 [07-slidelayout-fix.md](docs/07-slidelayout-fix.md)

### Case C：仅做 QA

1. Read [06-qa.md](docs/06-qa.md)
2. 跑 `scripts/qa_deep.py output.pptx` 看结构告警
3. 跑 `scripts/render_layout.py output.pptx qa_layout/` 生成布局缩略图，逐张人眼核查

## 依赖

```bash
pip install python-pptx>=1.0.2 Pillow
# 可选：LibreOffice (soffice) 用于 PPTX → PDF
```

详细依赖与版本兼容见各 doc 顶部说明。
