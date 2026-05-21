"""pptx-cjk 复用组件库。

所有组件都接受 slide + 位置参数,返回创建的 shape/chart 引用。
组件之间互不耦合,可任意组合。

模块组织:
- text: 文本框 / 多行 / 来源注释(基础)
- bullets: 红点列表
- kpi: KPI 卡片
- table: 浅灰表头表格(横线分隔,无外框)
- chart: 12 类原生 Chart(双击可编辑数据)
       column/bar/grouped_col/stacked_col/stacked_bar/pie/doughnut/
       line/area/theme_river/radar/scatter/bubble/stock
- diagrams: 7 类形状组合图(数据嵌入代码)
       funnel/process_flow/sankey/treemap/sunburst/gantt/heatmap
- hero: 标题页 / 分节扉页 / 难题扉页
- accent: 装饰元素(红锚、红 label、列分隔线、insight 卡、recap 块、bento)
- image: 安全图片插入(自动 aspect-fit + 居中)
"""

from . import text, bullets, kpi, table, chart, diagrams, hero, accent, image

__all__ = [
    "text", "bullets", "kpi", "table", "chart", "diagrams",
    "hero", "accent", "image",
]
