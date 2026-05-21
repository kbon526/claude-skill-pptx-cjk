# 10 — Asset Registry(资产注册表)

## 这是什么

**Asset Registry 是三管道工作流的中心化数据结构** —— A 内容管道和 B 视觉管道
的所有产出物,都以 `slide_id` 为 key 汇总在这个 JSON 里。C 设计管道再按
`slide_id` 一对一映射到具体的幻灯片。

可以把它看作"幻灯片的数据库",在生成 .pptx 之前,deck 已经在 JSON 里
"逻辑上存在"。

## 为什么需要它

### 问题:三管道并行后,如何汇合?
- A 管道产出文本/数据(JSON 友好)
- B 管道产出图片路径/图表 spec(JSON 友好)
- C 管道需要把两者按页对齐

### 传统做法的痛点
直接在 Python 内存里传 dict,每次改动都要重跑全流程:
```python
# 低效写法
slides_data = generate_content(...)
chart_data = generate_charts(...)
# 客户改了第 5 页文案 → 上面两步都要重跑
build_pptx(slides_data, chart_data)
```

### Asset Registry 写法
```python
# 高效写法
reg = init_registry(...)
add_slide(reg, "slide_05", title="...", copy={...})
attach_visual(reg, "slide_05", kind=KIND_NATIVE_CHART, spec={...})
save(reg, "work/registry.json")

# 客户改了第 5 页文案
reg = load("work/registry.json")
reg["slides"]["slide_05"]["copy"]["main"] = "新文案"
save(reg, "work/registry.json")
# 不用重跑 chart 生成,assemble 时自动用最新的 copy
```

## JSON Schema

```json
{
  "meta": {
    "deck_title": "AcmeCorp 2026 H2 媒介策略",
    "brand": "AcmeCorp",
    "version": "v1",
    "created_at": "2026-05-20T15:00:00",
    "checkpoint_status": {
      "ckpt_1_framework": "approved",
      "ckpt_2_content": "approved",
      "ckpt_3_visual": "pending",
      "ckpt_4_assembly": "pending"
    }
  },
  "slides": {
    "slide_01": {
      "title": "竞品 SOV 对比",
      "subtitle": "Q1-Q3 数据,Pathmatics",
      "layout": "two_column_kpi_bullets",
      "copy": {
        "main": "...",
        "bullets": ["GWM 占据 12% SOV", "Toyota 23% 领跑", "..."],
        "kpis": [
          {"val": "37%", "label": "Toyota Share"},
          {"val": "12%", "label": "GWM Share"}
        ]
      },
      "visuals": [
        {
          "kind": "native_chart",
          "spec": {
            "type": "doughnut",
            "cats": ["Toyota", "GWM", "Ford", "其他"],
            "vals": [37, 12, 18, 33],
            "colors": ["...", "...", "...", "..."]
          }
        },
        {
          "kind": "raster_png",
          "path": "work/images/hero_01.png",
          "prompt": "abstract competitive chart bg",
          "editable_via": "mcp_image_edit"
        }
      ],
      "notes": "演讲者备注..."
    }
  }
}
```

## Visual Kind 类型

| kind | 用途 | 来源管道 | 是否可编辑(in PPT) |
|---|---|---|---|
| `native_chart` | python-pptx 原生图表 | B-Data (B2) | ✅ 双击改数据 |
| `raster_png` | 栅格图(AI 生图) | B-Visual (B3) | ❌ 重新生成 |
| `svg_icon` | 矢量图标(Canva) | B-Visual (B4) | ✅ 解组改色 |
| `image_extracted` | 从素材提取的现成图 | B-Data (B1) | ❌ 重新提取 |
| `table` | 表格数据 | B-Data | ✅ 改单元格 |

## 核心 API

### 初始化
```python
from scripts.core.registry import init_registry
reg = init_registry("AcmeCorp 媒介策略", brand="AcmeCorp", version="v1")
```

### 添加 slide
```python
from scripts.core.registry import add_slide
add_slide(reg, "slide_01",
          title="竞品 SOV",
          subtitle="Q1-Q3 数据",
          layout="two_column_kpi_bullets",
          copy={"bullets": [...], "kpis": [...]})
```

### 附加视觉资产
```python
from scripts.core.registry import attach_visual, KIND_NATIVE_CHART, KIND_RASTER_PNG

# 原生图表
attach_visual(reg, "slide_01",
              kind=KIND_NATIVE_CHART,
              spec={"type": "doughnut", "cats": [...], "vals": [...], "colors": [...]})

# AI 生图
attach_visual(reg, "slide_01",
              kind=KIND_RASTER_PNG,
              path="work/images/hero_01.png",
              prompt="abstract bg",
              editable_via="mcp_image_edit")
```

### Checkpoint 状态
```python
from scripts.core.registry import set_checkpoint, all_approved, CKPT_APPROVED

set_checkpoint(reg, "ckpt_1_framework", CKPT_APPROVED)
if all_approved(reg):
    # 可以进入 C3 assemble
    pass
```

### 持久化
```python
from scripts.core.registry import save, load
save(reg, "work/registry.json")
reg = load("work/registry.json")
```

### 校验
```python
from scripts.core.registry import validate
errors = validate(reg)
if errors:
    print("Registry 有问题:")
    for e in errors:
        print(f"  - {e}")
```

### 摘要(用于 Checkpoint 提示)
```python
from scripts.core.registry import summary
print(summary(reg))
# →
# Deck: AcmeCorp 媒介策略 (AcmeCorp) v1
# Slides: 12
#   slide_01: '竞品 SOV' [layout=two_column_kpi_bullets, visuals=2 (native_chart, raster_png)]
#   ...
# Checkpoints:
#   ✅ ckpt_1_framework: approved
#   ⏸ ckpt_2_content: pending
```

## 设计原则

1. **slide_id 永久不变** — 一旦分配 `slide_01`,就一直是它,不要因为
   插入新页就 renumber。需要插入时用 `slide_05a` 或 `slide_05_5`。
2. **layout 字段决定 builder** — `assemble.py` 里的 `LAYOUT_BUILDERS` 字典
   做调度,新增布局只需写新 builder + 注册到字典。
3. **visuals 是 list 而非 dict** — 一页可以有多个图(主图 + 装饰图标),
   按附加顺序保留。
4. **Registry 是单一真相源** — 所有阶段读写同一个 JSON,避免多个版本同步问题。

## 何时不要用 Registry

如果只是"快速做一个 5 页小 deck",直接写 Python 脚本调 components 库即可,
不必动用 Registry —— 它的价值在于多管道协作和 checkpoint 管理。

5 页以内的简单 deck 直接用 [examples/01_minimal_deck.py](../examples/) 的写法。
