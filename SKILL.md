---
name: pptx-cjk
description: 中文 PPT 提案专精 skill，三管道工作流 + 4 强制 checkpoint + 强制模板。基于 python-pptx + 模板克隆路线，处理中文场景下的 PPTX 创建、修改与 QA。当用户要求"做一份提案 PPT/媒介策略 PPT/把 Word 转成 PPT/按某个模板重排 PPT/修改 PPTX 文件/中文幻灯片字体/微软雅黑/Hero 大字页/三色配色"等中文 PPT 任务时使用。也覆盖 python-pptx 操作、slideLayout XML 修复、占位符问题、PPTX 转图片做视觉 QA、CJK 字符宽度估算等场景。与官方 pptx skill 互补：官方走 pptxgenjs(JS) + markitdown 路线偏英文场景，本 skill 走 python-pptx + 模板克隆路线偏中文场景。
license: MIT
---

# pptx-cjk — 中文 PPT 提案专精 Skill

中文场景下的 PPTX 自动化制作与 QA。基于 python-pptx,沉淀自 BambuLab 媒介策略、Moloco CTV 提案、汽车竞品分析等多个真实项目。

---

## ⚠️ CRITICAL — 三大强约束

### 约束 1: 必须提供真实 PPTX 模板

**自 v0.3.0 起,`assemble()` 强制要求 `template_path` 参数,不再接受 None**。

如果你没有现成模板,先生成 starter:
```bash
python3 scripts/tools/generate_starter_template.py templates/starter.pptx
```

如果你想检查既有模板是否兼容:
```bash
python3 scripts/tools/validate_template.py path/to/your.pptx
```

详见 [docs/14-template-requirements.md](docs/14-template-requirements.md)。

### 约束 2: 4 个强制 Checkpoint

工作流强制在以下 4 个节点停下,等用户明确确认("继续"/"OK"/"go" 等)才能进入下一阶段:

| # | Checkpoint | 触发时机 | 产出物 | Claude 必须做的事 |
|---|---|---|---|---|
| 1 | **ckpt_1_framework** | A1 框架共创后 | `work/framework.md` | 调用 `ckpt_1_framework.announce()` 然后停下问用户 |
| 2 | **ckpt_2_content** | A3 内容映射后 | `work/registry.json`(slides 已填) | 调用 `ckpt_2_content.announce()` 然后停下问用户 |
| 3 | **ckpt_3_visual** | B4 视觉资产备齐后 | `work/registry.json`(visuals 已填) | 调用 `ckpt_3_visual.announce()` 然后停下问用户 |
| 4 | **ckpt_4_assembly** | C3 .pptx 组装后 | `output/<deck>.pptx` + 缩略图 | 调用 `ckpt_4_assembly.announce()` 然后停下问用户 |

**不允许"一气呵成"**(用户在 BambuLab V4→V5 经验中证明:大改稿成本远高于多轮小确认)。
若用户提出修改意见,在当前阶段内迭代,不要跳过本 checkpoint。

详见 [docs/11-checkpoints.md](docs/11-checkpoints.md)。

### 约束 3: 三铁律

1. **macOS 用户必须显式使用 Python 3.13 路径**,系统默认的 `python3` 大概率不含 `python-pptx`
   ```bash
   /Library/Frameworks/Python.framework/Versions/3.13/bin/python3 build.py
   ```
   详见 [docs/03-python-pptx-gotchas.md](docs/03-python-pptx-gotchas.md)

2. **脚本文件名严禁与 Python 标准库冲突**(特别是 `inspect.py` / `types.py` / `io.py`),会触发 lxml 循环导入

3. **中文字体严格使用微软雅黑 Light**(禁用 Regular,字距过宽)
   通过 `scripts/core/font.py` 的 `_sf()` 封装统一注入,加上 `patch_theme_fonts(prs)` 改主题字体。
   详见 [docs/04-cjk-fonts.md](docs/04-cjk-fonts.md)

---

## 何时使用本 Skill

✅ 用 Word 文档 / Markdown 内容生成中文 PPT 提案
✅ 按某个既有 PPTX 的版式克隆出新 PPT(保留品牌风格)
✅ 修改既有 PPTX(删页/插页/改标题/换字体/重排形状)
✅ 中文字体出现"显示成宋体""字号不对"等问题需要修复
✅ 客户交付前需要做 PPTX 的结构与视觉双轨 QA

❌ 英文场景从零创建 → 优先用官方 `pptx` skill
❌ 仅需要把 PPTX 转成 PDF → `soffice --headless --convert-to pdf` 即可,不必加载本 skill

---

## 三管道工作流(必读)

```
[输入: brief + 素材]
        │
   ┌────┴────┐
   ▼         ▼
[A 内容管道]  [B 视觉管道]
   │            │
   │     ┌──────┴──────┐
   │     │              │
   │   B-Data         B-Visual
   │   代码图表       AI 生图
   │     ↓              ↓
   │   ckpt_3 视觉确认 ←─┘
ckpt_2 内容确认       │
   ↓                    │
   └─────┬──────────────┘
         ▼
   [Asset Registry]  ← 中心化资产汇合
   slide_id → {copy, visuals}
         ↓
   [C 设计管道]  ← ⚠️ 必须提供 template_path
   C1 色板 → C2 模板 → C3 组装
         ↓
   ckpt_4 组装确认
         ↓
   [输出: .pptx]
```

详见 [docs/09-pipeline-architecture.md](docs/09-pipeline-architecture.md)。

---

## 标准工作流(分 5 阶段,严格按 Checkpoint 节奏)

### 阶段 0 — 准备模板(强约束前置步骤)
```bash
# 没有模板时,生成 starter
python3 scripts/tools/generate_starter_template.py templates/starter.pptx

# 有客户模板时,先校验
python3 scripts/tools/validate_template.py path/to/client_template.pptx
```

### 阶段 1 — A1 框架共创
- 与用户对齐 deck 标题、受众、章节结构、核心论点
- 写入 `work/framework.md`
- **🛑 ckpt_1_framework**

### 阶段 2 — A2/A3 内容摄入与映射
- A2:`scripts/pipeline/content_ingest.py` 读 Word/PDF/PPTX → 统一 markdown
- A3:`scripts/pipeline/content_allocate.py` 切片到 slide → registry.json 草稿
- **🛑 ckpt_2_content**

### 阶段 3 — B 视觉管道(并行)
**B-Data 路径(数据可视化,代码生成,可编辑)**:
- B1:`scripts/pipeline/visual_extract.py` 从素材提取已有图(可选)
- B2:`scripts/pipeline/chart_remake.py` 把数据 + 品牌色 → chart spec
- 详见 [docs/12-chart-factory.md](docs/12-chart-factory.md)

**B-Visual 路径(概念图/Hero 图)**:
- B3:`scripts/pipeline/image_provider.py` 调 ManualImageProvider / LocalAssetProvider /(Phase 2)MCPImageProvider
- B4:从 `assets/icons/` 选 svg 图标
- 详见 [docs/13-image-provider.md](docs/13-image-provider.md)

- **🛑 ckpt_3_visual**

### 阶段 4 — C 设计管道:组装(强制模板)
```python
from scripts.pipeline.assemble import assemble
assemble(reg,
         template_path="templates/starter.pptx",  # ⚠️ 必填!
         output_path="output/your_deck.pptx")
```
- **🛑 ckpt_4_assembly**

### 阶段 5 — 交付
- 把 `output/<deck>.pptx` 提交给用户
- 视觉 QA(可选):`scripts/qa_deep.py` + `scripts/render_layout.py`

---

## Quick Reference

### 概念性文档(理解架构)
| 你想理解什么 | 看哪份文档 |
|------------|------------|
| 三管道架构总览 | [docs/09-pipeline-architecture.md](docs/09-pipeline-architecture.md) |
| Asset Registry JSON schema | [docs/10-asset-registry.md](docs/10-asset-registry.md) |
| 4 个 Checkpoint 强约束 | [docs/11-checkpoints.md](docs/11-checkpoints.md) |
| 数据可视化代码生成路径 | [docs/12-chart-factory.md](docs/12-chart-factory.md) |
| AI 生图抽象层 | [docs/13-image-provider.md](docs/13-image-provider.md) |
| 模板要求 + 兼容性指南 | [docs/14-template-requirements.md](docs/14-template-requirements.md) |

### 操作性文档(实操避坑)
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

### 工具与示例
| 资源 | 路径 |
|---|---|
| 5 页最简示例 | [examples/01_minimal_deck.py](examples/01_minimal_deck.py) |
| 完整三管道示例 | [examples/02_brand_strategy.py](examples/02_brand_strategy.py) |
| 模板生成 | `scripts/tools/generate_starter_template.py` |
| 模板兼容性校验 | `scripts/tools/validate_template.py` |
| 模板/解包工具 | `scripts/inspect_template.py` / `scripts/extract_docx.py` |
| QA 工具 | `scripts/qa_deep.py` / `scripts/render_layout.py` |
| 文本工具 | `scripts/set_text.py` / `scripts/insert_slide_at.py` |

---

## 模块结构

```
scripts/
├── core/             # 核心基建
│   ├── font.py           # _sf() 字体注入 + 字号体系
│   ├── color.py          # RGB 常量池 + 品牌色字典
│   ├── layout.py         # 16:9 网格常数
│   └── registry.py       # Asset Registry IO + 校验
│
├── components/       # 复用组件库
│   ├── text.py / bullets.py / kpi.py
│   ├── table.py / chart.py
│   ├── hero.py / accent.py / image.py
│
├── pipeline/         # 三管道编排
│   ├── content_ingest.py     # A2 摄入
│   ├── content_allocate.py   # A3 映射
│   ├── visual_extract.py     # B1 提取
│   ├── chart_remake.py       # B2 图表重制
│   ├── image_provider.py     # B3 生图抽象层(Manual/Local/MCP)
│   └── assemble.py           # C3 组装(template_path 必填)
│
├── checkpoints/      # 4 个强制 gate
│   ├── _gate.py
│   └── ckpt_1_framework.py / ckpt_2_content.py / ckpt_3_visual.py / ckpt_4_assembly.py
│
├── tools/            # 命令行工具
│   ├── generate_starter_template.py   # 生成兼容模板
│   └── validate_template.py            # 校验模板兼容性
│
└── (顶层工具)        # 通用工具(继承自骨架)
    ├── inspect_template.py
    ├── extract_docx.py
    ├── insert_slide_at.py
    ├── set_text.py
    ├── qa_deep.py
    └── render_layout.py

templates/
└── starter.pptx      # 默认 starter 模板
```

---

## 不同任务的入口建议

### Case A:从零做提案(三管道完整工作流)
1. 准备模板:用 starter 或客户模板,跑 `validate_template.py`
2. Read [docs/09-pipeline-architecture.md](docs/09-pipeline-architecture.md)
3. 与用户共创框架 → **🛑 ckpt_1**
4. `scripts/pipeline/content_ingest.py` + `content_allocate.py`
5. **🛑 ckpt_2**
6. `scripts/pipeline/chart_remake.py` + `image_provider.py`
7. **🛑 ckpt_3**
8. `scripts/pipeline/assemble.py`(必传 template_path)
9. **🛑 ckpt_4**
10. 交付

参照 [examples/02_brand_strategy.py](examples/02_brand_strategy.py)。

### Case B:修改既有 PPT(轻量任务,跳过完整工作流)
1. Read [docs/03-python-pptx-gotchas.md](docs/03-python-pptx-gotchas.md) 避坑
2. Read [docs/08-shape-rebuild.md](docs/08-shape-rebuild.md) 理解形状操作模式
3. 如果要在中间插入 slide,用 `scripts/insert_slide_at.py`
4. 如果是标题被截断/字号问题,看 [docs/07-slidelayout-fix.md](docs/07-slidelayout-fix.md)

### Case C:仅做 QA
1. Read [docs/06-qa.md](docs/06-qa.md)
2. 跑 `scripts/qa_deep.py output.pptx` 看结构告警
3. 跑 `scripts/render_layout.py output.pptx qa_layout/` 生成布局缩略图

---

## 依赖

```bash
pip install python-pptx>=1.0.2 Pillow lxml
# 可选(B 管道用):
pip install pdfplumber python-docx markitdown
# 可选(QA):
# brew install libreoffice  # 提供 soffice 命令做 PPTX → PDF
```

详细依赖与版本兼容见各 doc 顶部说明。

---

## Phase 1 / Phase 2 路线

**Phase 1(当前 v0.3.0)**:
- ✅ 完整三管道架构 + 4 强制 Checkpoint
- ✅ 强制模板路线(template_path 必填)
- ✅ Blank layout + 自定义画位的稳定渲染
- ✅ 30+ 组件库(从 BambuLab/Moloco 项目沉淀)
- ✅ ManualImageProvider / LocalAssetProvider 可用
- ✅ 中国本土品牌色字典

**Phase 2(独立 mcp-image-gen 仓库)**:
- 🔜 包装 OpenAI Image API,重点支持 `image_edit`(局部精修)
- 🔜 替换 `MCPImageProvider` 实现
- 🔜 skill 调用代码无需修改,只切换 `get_provider("mcp")` 即可
