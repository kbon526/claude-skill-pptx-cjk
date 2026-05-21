"""B2 图表重制 — 把已知数据 + 品牌色 → 原生 chart spec / diagram spec。

应用场景:
- 客户给的素材是 PNG 截图(无源数据),需要重新画
- 客户给的是 Excel 数据,需要做品牌色版本
- 同一组数据要做多种风格(主红 / 蓝绿 / 暗色背景)

输出:写入 Asset Registry 的 spec(kind=native_chart 或 diagram)。
由 C3 assemble 时实例化为图表。

⚠️ 不在这里直接画图,而是产出 spec —— 这样三管道完全解耦,B2 重画时
不需要重新打开 .pptx。

API 命名约定:
- make_*_spec(...) 返回 dict
- attach_chart_to_slide(reg, slide_id, spec) 绑到 registry,自动判断 kind
"""

from typing import Any

from ..core.color import (
    RED, BLUE, GREEN, ORANGE,
    CLR_FACEBOOK, CLR_INSTAGRAM, CLR_TIKTOK, CLR_YOUTUBE,
    CLR_XHS, CLR_DOUYIN, CLR_LINE, CLR_WECHAT, CLR_WEIBO,
    LIGHT_GRAY, BRAND_COLORS,
)
from ..core.registry import attach_visual, KIND_NATIVE_CHART


# ══════════════════════════════════════════════════════════
# 类型分类:决定 attach 时用 native_chart 还是 diagram
# ══════════════════════════════════════════════════════════
NATIVE_CHART_TYPES = {
    "column", "bar", "grouped_col", "stacked_col", "stacked_bar",
    "pie", "doughnut",
    "line", "area", "theme_river",
    "radar",
    "scatter", "bubble",
    # ⚠️ "stock" 不在这里 — python-pptx 1.x 不支持 STOCK_OHLC 写入
}

DIAGRAM_TYPES = {
    "funnel", "process_flow", "sankey",
    "treemap", "sunburst", "gantt", "heatmap",
    "candlestick",  # K 线图(用形状绘制)
}


def _kind_of(chart_type: str) -> str:
    """根据 type 判断 kind(写入 registry 时用)。"""
    if chart_type in NATIVE_CHART_TYPES:
        return "native_chart"
    if chart_type in DIAGRAM_TYPES:
        return "diagram"
    raise ValueError(f"Unknown chart type: {chart_type}")


# ══════════════════════════════════════════════════════════
# A 类原生图表 spec 工厂
# ══════════════════════════════════════════════════════════
def make_column_spec(cats, vals, colors=None, fmt="#,##0") -> dict:
    return {"type": "column", "cats": cats, "vals": vals,
            "colors": colors, "fmt": fmt}


def make_bar_spec(cats, vals, colors=None, fmt="#,##0") -> dict:
    return {"type": "bar", "cats": cats, "vals": vals,
            "colors": colors, "fmt": fmt}


def make_grouped_col_spec(cats, series: dict) -> dict:
    """series: {name: (values, color)}"""
    return {"type": "grouped_col", "cats": cats, "series": series}


def make_stacked_col_spec(cats, series_list, percent=False) -> dict:
    """series_list: [(name, values, color), ...]"""
    return {"type": "stacked_col", "cats": cats,
            "series_list": series_list, "percent": percent}


def make_stacked_bar_spec(cats, series_list, percent=True) -> dict:
    return {"type": "stacked_bar", "cats": cats,
            "series_list": series_list, "percent": percent}


def make_pie_spec(cats, vals, colors, fmt='0"%"') -> dict:
    return {"type": "pie", "cats": cats, "vals": vals,
            "colors": colors, "fmt": fmt}


def make_doughnut_spec(cats, vals, colors, fmt='0"%"') -> dict:
    return {"type": "doughnut", "cats": cats, "vals": vals,
            "colors": colors, "fmt": fmt}


def make_line_spec(cats, series: dict, fmt="#,##0",
                    smooth=False, markers=True) -> dict:
    """series: {name: (values, color)}"""
    return {"type": "line", "cats": cats, "series": series,
            "fmt": fmt, "smooth": smooth, "markers": markers}


def make_area_spec(cats, series: dict, stacked=False, percent=False) -> dict:
    return {"type": "area", "cats": cats, "series": series,
            "stacked": stacked, "percent": percent}


def make_theme_river_spec(cats, series: dict) -> dict:
    return {"type": "theme_river", "cats": cats, "series": series}


def make_radar_spec(cats, series: dict, filled=False) -> dict:
    return {"type": "radar", "cats": cats, "series": series,
            "filled": filled}


def make_scatter_spec(series: dict, colors=None) -> dict:
    """series: {name: [(x, y), ...]}"""
    return {"type": "scatter", "series": series, "colors": colors}


def make_bubble_spec(series: dict, colors=None, bubble_scale=100) -> dict:
    """series: {name: [(x, y, size), ...]}"""
    return {"type": "bubble", "series": series, "colors": colors,
            "bubble_scale": bubble_scale}


def make_stock_spec(dates, ohlc, up_color=None, down_color=None) -> dict:
    """K 线图 spec(走 diagrams.add_candlestick,基于形状绘制)。

    Args:
        ohlc: list[(open, high, low, close)]
    """
    return {"type": "candlestick", "dates": dates, "ohlc": ohlc,
            "up_color": up_color, "down_color": down_color}


# ══════════════════════════════════════════════════════════
# B 类形状组合 diagram spec 工厂
# ══════════════════════════════════════════════════════════
def make_funnel_spec(stages, values=None, colors=None,
                      show_values=True, show_pct=True) -> dict:
    return {"type": "funnel", "stages": stages, "values": values,
            "colors": colors, "show_values": show_values, "show_pct": show_pct}


def make_process_flow_spec(steps, colors=None, arrow=True) -> dict:
    """steps: list[str | (title, desc)]"""
    return {"type": "process_flow", "steps": steps,
            "colors": colors, "arrow": arrow}


def make_sankey_spec(left_nodes, right_nodes, flows,
                      left_colors=None, right_colors=None) -> dict:
    """
    left_nodes: list[(name, value)]
    right_nodes: list[(name, value)]
    flows: list[(left_idx, right_idx, value)]
    """
    return {"type": "sankey",
            "left_nodes": left_nodes, "right_nodes": right_nodes,
            "flows": flows,
            "left_colors": left_colors, "right_colors": right_colors}


def make_treemap_spec(items, colors=None) -> dict:
    """items: list[(name, value)]"""
    return {"type": "treemap", "items": items, "colors": colors}


def make_sunburst_spec(hierarchy, colors=None) -> dict:
    """hierarchy: dict[parent -> dict[child -> value]] 或 list[(parent, child, value)]"""
    return {"type": "sunburst", "hierarchy": hierarchy, "colors": colors}


def make_gantt_spec(tasks, color=None) -> dict:
    """tasks: list[(task_name, start_day, duration_days)]"""
    return {"type": "gantt", "tasks": tasks, "color": color}


def make_heatmap_spec(matrix, x_labels, y_labels,
                       cmap_low=None, cmap_high=None, fmt="{:.0f}") -> dict:
    return {"type": "heatmap", "matrix": matrix,
            "x_labels": x_labels, "y_labels": y_labels,
            "cmap_low": cmap_low, "cmap_high": cmap_high, "fmt": fmt}


# ══════════════════════════════════════════════════════════
# 高阶辅助:常见品牌图自动配色
# ══════════════════════════════════════════════════════════
def channel_mix_spec(channel_data: dict) -> dict:
    """渠道占比图 spec - 自动用平台官方色。

    Args:
        channel_data: {"Facebook": 43, "Instagram": 23, ...}
    """
    color_map = {
        "Facebook": CLR_FACEBOOK,
        "Instagram": CLR_INSTAGRAM,
        "TikTok": CLR_TIKTOK,
        "YouTube": CLR_YOUTUBE,
        "LINE": CLR_LINE,
        "WeChat": CLR_WECHAT, "微信": CLR_WECHAT,
        "Weibo": CLR_WEIBO, "微博": CLR_WEIBO,
        "小红书": CLR_XHS, "XHS": CLR_XHS,
        "抖音": CLR_DOUYIN, "Douyin": CLR_DOUYIN,
    }
    cats = list(channel_data.keys())
    vals = list(channel_data.values())
    colors = [color_map.get(c, LIGHT_GRAY) for c in cats]
    return make_doughnut_spec(cats, vals, colors)


def brand_compare_spec(brands: list[str], values: list[float],
                        chart_type: str = "column") -> dict:
    """跨品牌对比图 - 自动用品牌官方色。"""
    colors = []
    for b in brands:
        palette = BRAND_COLORS.get(b)
        colors.append(palette["primary"] if palette else BLUE)
    if chart_type == "bar":
        return make_bar_spec(brands, values, colors)
    return make_column_spec(brands, values, colors)


# ══════════════════════════════════════════════════════════
# 通用绑定方法
# ══════════════════════════════════════════════════════════
def attach_chart_to_slide(reg: dict, slide_id: str, spec: dict,
                           title: str = ""):
    """把 chart spec 绑到 registry,自动根据 type 决定 kind。

    - type ∈ NATIVE_CHART_TYPES → kind="native_chart"
    - type ∈ DIAGRAM_TYPES → kind="diagram"
    """
    chart_type = spec.get("type")
    kind = _kind_of(chart_type)
    attach_visual(reg, slide_id, kind=kind, spec=spec, title=title)
