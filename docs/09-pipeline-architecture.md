# 09 — 三管道架构总览

本 skill 的工作流采用**三管道并行 + 中心化资产汇合 + 4 段 Checkpoint** 的设计。
这是从多个真实中文商务 PPT 项目(BambuLab 媒介策略、Moloco CTV、汽车竞品分析)
沉淀而来的反返工模式。

## 总览图

```
[输入: brief + 素材]
        │
   ┌────┴────┐
   ▼         ▼
[A 内容管道]  [B 视觉管道]
   │         │
   │  ┌──────┴──────┐
   │  │              │
   │  ▼ B-Data       ▼ B-Visual
   │  代码生成图表    AI 生图/素材
   │   ↓              ↓
   ↓  ckpt_3 视觉确认 ←─┘
ckpt_2 内容确认       │
   ↓                  │
   └─────┬────────────┘
         ▼
   [Asset Registry]  ← 中心化资产汇合
   slide_id → {copy, visuals}
         ↓
   [C 设计管道]
   C1 色板 → C2 模板 → C3 组装
         ↓
   ckpt_4 组装确认
         ↓
   [输出: .pptx]
```

## 三条管道的职责

### A 内容管道(Content Pipeline)
- **A1 框架共创**:与用户对齐 deck 标题/受众/章节结构 → 产出 `framework.md`
  - 🛑 **ckpt_1_framework**:用户确认框架后才进入摄入
- **A2 素材摄入**:把 PDF/DOCX/PPTX/Markdown 统一转为 Markdown
  - 入口:`scripts/pipeline/content_ingest.py`
- **A3 内容映射**:把 Markdown 切片到每张 slide,识别 KPI/bullet/段落
  - 入口:`scripts/pipeline/content_allocate.py`
  - 输出:草稿 `registry.json`(`slides` 字段已填,`visuals` 为空)
  - 🛑 **ckpt_2_content**:用户确认每页内容后才进入视觉管道
- **A4 文案精修**:根据 ckpt_2 反馈,在内容层做最后润色

### B 视觉管道(Visual Pipeline)
B 管道**主动分流**为两条互不阻塞的子管道:

#### B-Data:数据可视化(代码生成)
- **B1 视觉提取**:从素材里抽出已有图片/图表(`visual_extract.py`)
- **B2 图表重制**:把数据 + 品牌色 → chart spec(`chart_remake.py`)
- **优势**:python-pptx 原生 Chart,客户在 PPT 里可双击直接编辑

#### B-Visual:概念图/装饰图(AI 生图)
- **B3 AI 生图**:通过 `image_provider` 抽象层调用
  - Phase 1:`ManualImageProvider`(手动产图后回填)
  - Phase 1:`LocalAssetProvider`(本地素材库匹配)
  - Phase 2:`MCPImageProvider`(调 mcp-image-gen,接 OpenAI Image API)
- **B4 资产采购**:Canva 图标/svg(可手动放入 `assets/`)

🛑 **ckpt_3_visual**:用户确认所有视觉资产后才进入设计管道

### C 设计管道(Design Pipeline)
- **C1 色板**:从 `core.color.BRAND_COLORS` 取品牌主色,注入图表/装饰元素
- **C2 模板匹配**:挑选最贴合的 .pptx 模板(模板克隆策略,见 02-template-clone)
- **C3 组装引擎**:`assemble.py` 根据 layout 字段调度对应 builder,
  生成最终 .pptx
- 🛑 **ckpt_4_assembly**:.pptx 生成完毕后做 QA + 用户最终确认

## 为什么是"管道"而不是"线性流程"

线性流程(A → B → C 串行)的问题:
- 客户改一处文案 → 整个流程从头跑
- 视觉返工要等内容定稿,机械等待

管道并行的好处:
- A 管道还在润色文案时,B 管道已经能开始处理数据可视化
- 改一处 → 只重跑那一段,Asset Registry 把局部变化隔离

## 为什么 B 要分流(B-Data / B-Visual)

详见 [10-asset-registry.md](10-asset-registry.md) 和 [13-image-provider.md](13-image-provider.md)。

核心原因:
- **B-Data 数据可视化必须可编辑**(客户要在 PPT 里改数据/颜色),
  所以走代码生成路径,产出原生 Chart 对象
- **B-Visual 概念图可以是栅格图**(改稿就重新生成),所以走 AI 生图路径

混在一起会带来"图表也用 AI 生图导致客户改不了"的灾难。

## 与 Checkpoint 的关系

详见 [11-checkpoints.md](11-checkpoints.md)。

简而言之:**4 个 Checkpoint 是契约式停顿点,Claude 在主对话中必须停下等用户确认**。
不允许"一气呵成"的工作流 —— 这是从 BambuLab V4→V5 16 处标题改写的痛苦经验里总结的规律。

## 入口脚本一览

| 阶段 | 脚本 | 输入 | 输出 |
|---|---|---|---|
| A2 摄入 | `pipeline/content_ingest.py` | PDF/DOCX/PPTX/MD | 统一 Markdown |
| A3 映射 | `pipeline/content_allocate.py` | Markdown | registry.json 草稿 |
| B1 提取 | `pipeline/visual_extract.py` | PDF/PPTX | extracted/*.png |
| B2 重制 | `pipeline/chart_remake.py` | 数据 + 品牌色 | chart spec |
| B3 生图 | `pipeline/image_provider.py` | prompt | image path |
| C3 组装 | `pipeline/assemble.py` | registry.json + template.pptx | output.pptx |
| Checkpoint | `checkpoints/ckpt_*.py` | registry | 提示文本 |
