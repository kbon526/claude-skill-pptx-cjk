# pptx-cjk

> Claude 中文 PPT 提案专精 Skill — **三管道工作流 + 4 强制 Checkpoint**,基于 python-pptx + 模板克隆路线,覆盖 Word→PPT、字体规范、AI 生图集成、双轨 QA、版式修复等全流程。

[![Skill version](https://img.shields.io/badge/skill-0.2.0-brightgreen)](#)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.10+-yellow.svg)](#)

## 这是什么

这是一个为 [Claude Code / Claude Agent SDK](https://docs.claude.com/) 编写的 Skill。
专注于**中文场景下的 PPT 提案制作**,沉淀了在多个真实提案项目(BambuLab 媒介策略、
Moloco CTV 投放方案、汽车竞品分析等)中验证过的工作流、踩过的坑和复用脚本。

### 核心特性

- 🎯 **三管道架构** — A 内容 / B 视觉 / C 设计三条并行管道,通过 Asset Registry 中心化汇合
- 🛑 **4 强制 Checkpoint** — 框架/内容/视觉/组装四个节点强制人工确认,反返工
- 🇨🇳 **中文优先** — 微软雅黑 Light 全栈注入,品牌色字典覆盖小米/比亚迪/华为/字节等
- 📊 **数据可视化代码生成** — python-pptx 原生 Chart,客户在 PPT 里可双击编辑
- 🎨 **AI 生图抽象层** — Phase 1 支持手动/本地素材库,Phase 2 接 mcp-image-gen
- 🧰 **沉淀级组件库** — KPI/bullets/table/chart/hero/accent 等 30+ helpers,从生产项目蒸馏

## 与官方 `pptx` skill 的差异

Anthropic 官方提供的 `pptx` skill 偏向英文场景、走 `markitdown + pptxgenjs(JS)` 路线。
本 skill 是**互补**而非替代:

| 维度 | 官方 `pptx` | `pptx-cjk` |
|------|-------------|-----------|
| 主要语言场景 | 英文 | 中文(CJK)|
| 创建技术栈 | pptxgenjs (JS) | python-pptx |
| 核心方法 | 从零创建 | **三管道 + 模板克隆** |
| 字体规范 | 通用排版 | 微软雅黑 Light 严格注入 |
| 工作流约束 | 灵活 | **4 个强制 Checkpoint** |
| 数据可视化 | 默认栅格图 | **原生 Chart(可编辑)** |
| AI 生图 | 不集成 | **image_provider 抽象层** |
| QA 方案 | LibreOffice → PDF | 双轨:结构(qa_deep) + 视觉(PIL render) |

两个 skill 可以共存。当你做中文提案、需要克隆既有 PPT 模板、或要把 Word 文档转成 PPT 时,Claude 会优先匹配本 skill。

## 安装

将本仓库克隆到 Claude 的用户级 skills 目录:

```bash
git clone https://github.com/kbon526/claude-skill-pptx-cjk.git ~/.claude/skills/pptx-cjk
```

或者作为项目级 skill(仅当前项目可用):

```bash
git clone https://github.com/kbon526/claude-skill-pptx-cjk.git .claude/skills/pptx-cjk
```

安装后下次启动 Claude Code 即可使用。

## 包含什么

```
pptx-cjk/
├── SKILL.md                       # Skill 主入口(三管道 + checkpoint 强约束)
├── docs/                          # 13 篇分主题深度文档
│   ├── 01-workflow.md                 # Word/DOCX → PPT 完整流程
│   ├── 02-template-clone.md           # 模板克隆构建法
│   ├── 03-python-pptx-gotchas.md      # python-pptx 关键坑位
│   ├── 04-cjk-fonts.md                # 中文字体规范
│   ├── 05-design-patterns.md          # 大字 Hero / 三色配色 / 标题精修
│   ├── 06-qa.md                       # 双轨 QA
│   ├── 07-slidelayout-fix.md          # slideLayout XML 修复
│   ├── 08-shape-rebuild.md            # 形状重建
│   ├── 09-pipeline-architecture.md    # ⭐ 三管道架构总览
│   ├── 10-asset-registry.md           # ⭐ Asset Registry JSON schema
│   ├── 11-checkpoints.md              # ⭐ 4 强制 Checkpoint
│   ├── 12-chart-factory.md            # ⭐ 数据可视化代码生成
│   └── 13-image-provider.md           # ⭐ AI 生图抽象层
├── scripts/
│   ├── core/                      # ⭐ 核心基建
│   │   ├── font.py                    # _sf() 字体注入 + 字号体系
│   │   ├── color.py                   # RGB 常量池 + BRAND_COLORS
│   │   ├── layout.py                  # 16:9 网格常数体系
│   │   └── registry.py                # Asset Registry IO
│   ├── components/                # ⭐ 复用组件库
│   │   ├── text.py / bullets.py / kpi.py
│   │   ├── table.py / chart.py
│   │   ├── hero.py / accent.py / image.py
│   ├── pipeline/                  # ⭐ 三管道编排脚本
│   │   ├── content_ingest.py / content_allocate.py
│   │   ├── visual_extract.py / chart_remake.py
│   │   ├── image_provider.py / assemble.py
│   ├── checkpoints/               # ⭐ 4 强制 gate
│   │   ├── _gate.py
│   │   └── ckpt_1_framework.py / ckpt_2_content.py / ckpt_3_visual.py / ckpt_4_assembly.py
│   ├── inspect_template.py        # 工具脚本(继承)
│   ├── insert_slide_at.py
│   ├── extract_docx.py
│   ├── qa_deep.py
│   ├── render_layout.py
│   └── set_text.py
├── examples/
│   ├── 01_minimal_deck.py             # 5 页中文 deck 最简示例
│   └── 02_brand_strategy.py           # 完整三管道 + checkpoint 实战示例
└── tests/
    ├── test_core_font.py
    ├── test_core_color.py
    └── test_components.py
```

## 使用要求

- Python 3.10+ 推荐(macOS 默认 python3 可能不含 python-pptx,详见 [python-pptx-gotchas.md](docs/03-python-pptx-gotchas.md))
- 必需:`python-pptx >= 1.0.2`、`Pillow`、`lxml`
- 可选:`pdfplumber`、`python-docx`、`markitdown`(B 管道摄入素材用)
- 可选:LibreOffice(`soffice`)用于 PPTX → PDF 转换

```bash
pip install python-pptx Pillow lxml
# B 管道增强(摄入 PDF/DOCX/PPTX 素材)
pip install pdfplumber python-docx markitdown
```

## 快速开始

### 完整三管道工作流
```
> 帮我做一份给小米生态产品的中文媒介策略 deck,brief 在 ~/Downloads/小米brief.docx
```

Claude 会:
1. 读取 brief,共创 deck 框架 → **🛑 ckpt_1**(等你确认)
2. 摄入素材,映射到每页 → **🛑 ckpt_2**(等你确认)
3. 并行生成图表(代码) + 概念图(image_provider) → **🛑 ckpt_3**(等你确认)
4. 组装 .pptx + QA → **🛑 ckpt_4**(等你确认)
5. 交付最终文件

### 轻量任务
```
> 把 ~/Downloads/方案.docx 按照 ~/Desktop/参考模板.pptx 的版式做成中文 PPT
```

Claude 会按照 [docs/01-workflow.md](docs/01-workflow.md) 的流程执行。

## Phase 路线

### Phase 1(当前 v0.2.0)
- ✅ 完整三管道架构 + 4 强制 Checkpoint
- ✅ 30+ 组件库(从 BambuLab/Moloco 项目沉淀)
- ✅ ManualImageProvider + LocalAssetProvider
- ✅ 中国本土品牌色字典

### Phase 2(规划中)
- 🔜 独立仓库 [mcp-image-gen](https://github.com/kbon526/mcp-image-gen)(包装 OpenAI Image API)
- 🔜 MCPImageProvider 落地,支持 `image_edit` 局部精修
- 🔜 更多 examples(产品发布会、品牌年报)
- 🔜 GitHub Actions(自动化测试)

## 贡献

如果你在使用中遇到 python-pptx 的新坑或验证了新的 pattern,欢迎提 PR。建议遵循:

- 每个 doc 聚焦一个主题,加 **场景** / **解法** / **注意** 三段结构
- 脚本必须自包含、可独立运行(不依赖业务数据)
- 新增品牌色到 `core/color.py BRAND_COLORS`(欢迎补充)
- 新增组件到 `components/`,在 `__init__.py` 注册

## 许可证

MIT © 2026 zz
