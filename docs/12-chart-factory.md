# 12 — Chart Factory(18 种图表 + 决策方法论)

## 方法论:两步法

### 第一步:数据类型 → 图表类型(决策树)

根据**数据想表达什么**选择图表:

| 数据意图 | 适用图表 | 推荐 top 1 |
|---|---|---|
| **比较**(不同实体的同一指标) | 柱/条/分组柱/雷达 | 柱状图 |
| **占比**(整体的组成) | 饼/环形/堆叠柱(100%)/树图/旭日 | 环形图(≤5 类) / 树图(>10 类) |
| **趋势**(时间序列变化) | 折线/面积/主题河流/K线 | 折线图 |
| **关系**(多变量相关性) | 散点/气泡/热力图 | 散点图 |
| **分布**(数据稠密度) | 热力图/散点 | 热力图 |
| **流向**(A → B 的流转) | 桑基/漏斗/流程图 | 漏斗图(单向衰减) |
| **层级**(嵌套结构) | 树图/旭日 | 树图 |
| **排程**(任务/里程碑) | 甘特图 | 甘特图 |

**自动推荐**:
```python
from scripts.pipeline.chart_picker import auto_select, explain

data = {"cats": ["1月", "2月", "3月"], "series": {"GMV": ([100, 120, 140], None)}}
recs = auto_select(data, intent="trend")
print(explain(recs))
# →
# 1. 🟢 line  (95%)  理由:单系列时间序列,折线图首选
# 2. 🟢 area  (70%)  理由:强调累计量时用面积图
# 3. 🟢 column  (60%)  理由:类目少且想强调离散值,柱状图也可
```

### 第二步:生成可在 PPT 中编辑的图表

把推荐结果转为 chart spec,写入 Asset Registry,由 C3 assemble 生成。

```python
from scripts.pipeline.chart_remake import make_line_spec, attach_chart_to_slide

spec = make_line_spec(
    cats=["1月", "2月", "3月"],
    series={"GMV": ([100, 120, 140], RED)},
)
attach_chart_to_slide(reg, "slide_03", spec, title="GMV 趋势")
```

assemble 时自动调用 `chart.render_from_spec` 或 `diagrams.render_from_spec`(根据 type 分流)。

---

## 18 种图表的实现矩阵

### 🟢 A 类:原生 Chart(双击可编辑数据)

| # | 类型 | type 字段 | XL_CHART_TYPE | 工厂函数 |
|---|---|---|---|---|
| 1 | 柱状图 | `column` | COLUMN_CLUSTERED | `make_column_spec` |
| 2 | 条形图 | `bar` | BAR_CLUSTERED | `make_bar_spec` |
| 3 | 分组柱状图 | `grouped_col` | COLUMN_CLUSTERED(多系列) | `make_grouped_col_spec` |
| 4 | 堆叠柱状图 | `stacked_col` | COLUMN_STACKED / _100 | `make_stacked_col_spec` |
| 5 | 堆叠条形图 | `stacked_bar` | BAR_STACKED / _100 | `make_stacked_bar_spec` |
| 6 | 饼图 | `pie` | PIE | `make_pie_spec` |
| 7 | 环形图 | `doughnut` | DOUGHNUT | `make_doughnut_spec` |
| 8 | 折线图 | `line` | LINE | `make_line_spec` |
| 9 | 面积图 | `area` | AREA / AREA_STACKED / _100 | `make_area_spec` |
| 10 | 主题河流图 | `theme_river` | AREA_STACKED 模拟 | `make_theme_river_spec` |
| 11 | 雷达图 | `radar` | RADAR_MARKERS / RADAR_FILLED | `make_radar_spec` |
| 12 | 散点图 | `scatter` | XY_SCATTER | `make_scatter_spec` |
| 13 | 气泡图 | `bubble` | BUBBLE | `make_bubble_spec` |

### 🟡 B 类:Diagram(形状组合,数据嵌入代码)

| # | 类型 | type 字段 | 实现方式 | 工厂函数 |
|---|---|---|---|---|
| 14 | K 线图 | `candlestick` | 形状组合(影线 + 实体矩形) | `make_stock_spec` |
| 15 | 漏斗图 | `funnel` | 矩形堆叠(简化梯形) | `make_funnel_spec` |
| 16 | 流程图 | `process_flow` | 圆角矩形 + 右箭头 | `make_process_flow_spec` |
| 17 | 桑基图 | `sankey` | 矩形 + 直线连接(简化版) | `make_sankey_spec` |
| 18 | 矩形树图 | `treemap` | 嵌套矩形(squarify 算法) | `make_treemap_spec` |
| 19 | 旭日图 | `sunburst` | 同心环形图(双层简化) | `make_sunburst_spec` |
| 20 | 甘特图 | `gantt` | BAR_STACKED 第一系列透明 | `make_gantt_spec` |
| 21 | 热力图 | `heatmap` | 表格 + 渐变填充 | `make_heatmap_spec` |

**为什么有些图必须走 diagram**:
- python-pptx 不支持的类型:STOCK_OHLC、SUNBURST、TREEMAP、FUNNEL
- 不存在的 XL_CHART_TYPE:桑基图、流程图、热力图

**为什么甘特图能做但归到 B 类**:
- BAR_STACKED 第一系列设为透明 = 起始偏移技巧
- 数据可在 Excel 里改,但理解成本高,标记为 B 类是为了让客户改稿时调用我们而非自己改

---

## 完整调用流程(从决策到生成)

```python
from scripts.pipeline.chart_picker import auto_select
from scripts.pipeline.chart_remake import (
    make_doughnut_spec, attach_chart_to_slide,
)
from scripts.core.color import RED, BLUE, GREEN, ORANGE

# 1. 数据
data = {
    "cats": ["品牌 A", "品牌 B", "品牌 C", "品牌 D"],
    "vals": [38, 23, 24, 15],  # 总和接近 100 → 自动识别为占比
}

# 2. 自动推荐
recs = auto_select(data)
# → top 1: doughnut(90%)

# 3. 用推荐的类型构造 spec
spec = make_doughnut_spec(
    cats=data["cats"],
    vals=data["vals"],
    colors=[RED, BLUE, GREEN, ORANGE],
)

# 4. 绑到 registry(自动判断 kind=native_chart)
attach_chart_to_slide(reg, "slide_07", spec)

# 5. assemble 时由 chart.render_from_spec 渲染
# 客户在 PPT 里双击图表 → Excel 数据表打开,可改数字
```

---

## 高阶辅助方法

### 渠道分布自动配色
```python
from scripts.pipeline.chart_remake import channel_mix_spec

spec = channel_mix_spec({
    "Facebook": 43, "Instagram": 23,
    "TikTok": 19, "YouTube": 15,
    "微信": 22, "小红书": 28, "抖音": 32,
})
# 自动用平台官方色(Facebook 蓝、Instagram 粉、抖音黑、微信绿等)
```

### 品牌对比自动配色
```python
from scripts.pipeline.chart_remake import brand_compare_spec

spec = brand_compare_spec(
    brands=["小米", "华为", "vivo", "OPPO"],
    values=[28, 32, 18, 22],
    chart_type="column",
)
# 自动从 BRAND_COLORS 字典取官方品牌色
```

---

## 客户改稿场景

### A 类(原生 Chart)
- **改数据** → PPT 里双击图表 → Excel 弹出 → 改数字 → 关闭即生效
- **改色** → 选中数据点 → 右键 "填充" → 选色
- **改图例位置** → 选中图例 → 拖动

### B 类(Diagram)
- **改形状色/位置** → PPT 里直接拖动选择形状,改色
- **改数据** → 必须回到代码改 spec,重新跑 assemble

提示:Asset Registry 里的 `editable_via` 字段标注修改方式:
```json
{
  "kind": "diagram",
  "spec": {"type": "funnel", ...},
  "editable_via": "code_only"
}
```

---

## 不支持的图表类型

| 图表 | 原因 | 替代方案 |
|---|---|---|
| **词云** | 视觉价值低,Python 库 `wordcloud` 生成 PNG 嵌入更好 | 用 `wordcloud` + image_provider |
| **3D 图表** | 商务 PPT 反模式(数据失真) | 用 2D 替代 |
| **真正的多层旭日图** | python-pptx 限制 | 提供双层简化版,或用外部库生成图嵌入 |

---

## 实战速查表

**给我**     →  **用什么图**
- 时间序列 N 个月的 GMV → 折线图
- 4 个品牌的 SOV 占比 → 环形图(≤5 类) / 堆叠条(>5 类)
- 3 个品牌 × 5 个渠道 → 100% 堆叠条形图
- 5 个维度的品牌评分对比 → 雷达图
- 营销漏斗(曝光→点击→下单) → 漏斗图
- 客户旅程 6 个步骤 → 流程图
- 渠道流量去向(N 源 → M 目标) → 桑基图
- 品类占比(>10 个品类) → 矩形树图
- 任务排程 + 时间偏移 → 甘特图
- 时段 × 维度的热度矩阵 → 热力图
- 股票/汇率走势 → K 线图

完整决策树见 [scripts/pipeline/chart_picker.py](../scripts/pipeline/chart_picker.py)。
