# pptx-cjk

> Claude 中文 PPT 提案专精 Skill — 基于 python-pptx + 模板克隆路线，覆盖 Word→PPT、字体规范、双轨 QA、版式修复等全流程。

[![Skill version](https://img.shields.io/badge/skill-0.1.0-brightgreen)](#)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

## 这是什么

这是一个为 [Claude Code / Claude Agent SDK](https://docs.claude.com/) 编写的 Skill。专注于**中文场景下的 PPT 提案制作**，沉淀了在多个真实提案项目（媒介策略、CTV 投放方案等）中验证过的工作流、踩过的坑和复用脚本。

## 与官方 `pptx` skill 的差异

Anthropic 官方提供的 `pptx` skill 偏向英文场景、走 `markitdown + pptxgenjs(JS)` 路线。本 skill 是**互补**而非替代：

| 维度 | 官方 `pptx` | `pptx-cjk` |
|------|-------------|-----------|
| 主要语言场景 | 英文 | 中文（CJK）|
| 创建技术栈 | pptxgenjs (JS) | python-pptx |
| 核心方法 | 从零创建 | 模板克隆 + 占位符复用 |
| 字体规范 | 通用排版 | 微软雅黑 Bold/Light 严格区分 |
| QA 方案 | LibreOffice → PDF → 视觉 | 双轨：结构(qa_deep) + 视觉(PIL render) |
| 输入源 | 直接生成 | 含 DOCX → PPT 提取流程 |

两个 skill 可以共存。当你做中文提案、需要克隆既有 PPT 模板、或要把 Word 文档转成 PPT 时，Claude 会优先匹配本 skill。

## 安装

将本仓库克隆到 Claude 的用户级 skills 目录：

```bash
git clone https://github.com/kbon526/claude-skill-pptx-cjk.git ~/.claude/skills/pptx-cjk
```

或者作为项目级 skill（仅当前项目可用）：

```bash
git clone https://github.com/kbon526/claude-skill-pptx-cjk.git .claude/skills/pptx-cjk
```

安装后下次启动 Claude Code 即可使用。

## 包含什么

```
pptx-cjk/
├── SKILL.md                 # Skill 主入口（Claude 加载它）
├── docs/                    # 8 篇分主题深度文档
│   ├── 01-workflow.md           # Word/DOCX → PPT 完整流程
│   ├── 02-template-clone.md     # 模板克隆构建法
│   ├── 03-python-pptx-gotchas.md  # python-pptx 关键坑位
│   ├── 04-cjk-fonts.md          # 中文字体规范
│   ├── 05-design-patterns.md    # 大字 Hero / 三色配色 / 标题精修
│   ├── 06-qa.md                 # 双轨 QA
│   ├── 07-slidelayout-fix.md    # slideLayout XML 修复
│   └── 08-shape-rebuild.md      # 形状重建
├── scripts/                 # 可复用工具脚本
│   ├── inspect_template.py      # 解包 + 枚举 layouts + dump 占位符
│   ├── insert_slide_at.py       # 任意位置插入 slide 的 helper
│   ├── extract_docx.py          # 直接解压 DOCX 提取 XML（零依赖）
│   ├── qa_deep.py               # 结构 QA
│   ├── render_layout.py         # PIL 渲染 shape 几何边界
│   └── set_text.py              # 字体规范封装
└── examples/
    └── minimal_deck.py      # 5 页中文 deck 最小示例
```

## 使用要求

- Python 3.10+ 推荐（macOS 默认 python3 可能不含 python-pptx，详见 [python-pptx-gotchas.md](docs/03-python-pptx-gotchas.md)）
- `python-pptx >= 1.0.2`
- `Pillow`（仅 `render_layout.py` 需要）
- 可选：LibreOffice（`soffice`）用于 PPTX → PDF 转换

```bash
pip install python-pptx Pillow
```

## 快速开始

进入 Claude Code，直接说：

> 把 `Downloads/方案.docx` 按照 `Desktop/参考模板.pptx` 的版式做成中文 PPT 提案

Claude 会自动加载本 skill，按照 [docs/01-workflow.md](docs/01-workflow.md) 的流程执行。

## 贡献

如果你在使用中遇到 python-pptx 的新坑或验证了新的 pattern，欢迎提 PR。建议遵循：

- 每个 doc 聚焦一个主题，加 **场景** / **解法** / **注意** 三段结构
- 脚本必须自包含、可独立运行（不依赖业务数据）

## License

MIT © 2026 zz
