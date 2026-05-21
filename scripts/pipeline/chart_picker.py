"""chart_picker.py — 数据类型 → 图表选型决策树。

核心目标:给定数据 + 表达意图,推荐 top 3 候选图表类型 + 选型理由。

输入约定(任选其一):
1. 显式 intent + data:
    intent: 比较 / 占比 / 趋势 / 关系 / 分布 / 流向 / 层级 / 排程
    data: 见各 intent 的格式说明

2. 仅 data,自动识别 intent(基于数据结构特征)

输出格式:
    [
      {"type": "column", "kind": "native_chart",
       "reason": "...", "confidence": 0.9,
       "spec_template": {...}},
      ...
    ]

设计参考:
- Andrew Abela 的 Chart Chooser
- Tableau "Show Me" 推荐器
- Holtz 的 "From Data to Viz"
"""

from typing import Literal


# 意图分类(对齐 docs/12-chart-factory.md 决策树)
Intent = Literal[
    "comparison",     # 比较:不同实体的同一指标
    "composition",    # 占比:整体的组成
    "trend",          # 趋势:时间序列变化
    "relationship",   # 关系:多变量相关性
    "distribution",   # 分布:数据稠密度
    "flow",           # 流向:A → B 的流转
    "hierarchy",      # 层级:嵌套结构
    "schedule",       # 排程:任务/里程碑
]

# 中文同义词映射
INTENT_ALIASES = {
    "比较": "comparison", "对比": "comparison",
    "占比": "composition", "构成": "composition", "分布": "composition",
    "趋势": "trend", "变化": "trend", "时间序列": "trend",
    "关系": "relationship", "相关": "relationship",
    "稠密": "distribution",
    "流向": "flow", "流转": "flow", "漏斗": "flow", "转化": "flow",
    "层级": "hierarchy", "嵌套": "hierarchy",
    "排程": "schedule", "甘特": "schedule", "里程碑": "schedule",
}


def normalize_intent(intent: str | None) -> str | None:
    """中文意图 → 英文标准化。"""
    if intent is None:
        return None
    if intent in INTENT_ALIASES:
        return INTENT_ALIASES[intent]
    if intent in (
        "comparison", "composition", "trend",
        "relationship", "distribution", "flow",
        "hierarchy", "schedule",
    ):
        return intent
    return None


# ══════════════════════════════════════════════════════════
# 数据特征识别
# ══════════════════════════════════════════════════════════
def analyze_data(data: dict) -> dict:
    """识别数据特征,返回结构化描述。

    Args:
        data: 灵活格式,典型字段:
            cats: 类目列表
            vals / values: 数值列表(单系列)
            series: 多系列数据
            time_axis: 是否含时间维度
            hierarchy: 是否含层级嵌套
            ...

    Returns:
        {
            "n_cats": 类目数,
            "n_series": 系列数,
            "is_time": 是否时间序列,
            "is_hierarchy": 是否层级数据,
            "is_flow": 是否流向数据,
            "is_paired": 是否成对(x, y) 散点,
            "value_sums_to_100": 占比总和 ≈ 100%,
        }
    """
    feat = {
        "n_cats": 0, "n_series": 0,
        "is_time": False, "is_hierarchy": False,
        "is_flow": False, "is_paired": False,
        "value_sums_to_100": False,
    }

    cats = data.get("cats", []) or data.get("categories", [])
    feat["n_cats"] = len(cats)

    # 时间序列识别(类目里含日期/月/季度等关键词)
    time_kw = ["年", "月", "日", "季度", "Q1", "Q2", "Q3", "Q4",
               "20", "19", "Jan", "Feb", "周"]
    if cats:
        sample = str(cats[0])
        feat["is_time"] = any(k in sample for k in time_kw)

    # 多系列
    if "series" in data:
        feat["n_series"] = len(data["series"])
    elif "vals" in data or "values" in data:
        feat["n_series"] = 1

    # 占比和(若有 vals 总和 ≈ 100,提示占比)
    vals = data.get("vals") or data.get("values")
    if vals and isinstance(vals, list):
        try:
            total = sum(vals)
            if 95 <= total <= 105:
                feat["value_sums_to_100"] = True
        except TypeError:
            pass

    # 层级
    if "hierarchy" in data or "tree" in data:
        feat["is_hierarchy"] = True

    # 流向
    if "flows" in data or ("left_nodes" in data and "right_nodes" in data):
        feat["is_flow"] = True

    # 成对散点
    if "series" in data and isinstance(data["series"], dict):
        first_v = next(iter(data["series"].values()), None)
        if isinstance(first_v, list) and first_v and isinstance(first_v[0], tuple):
            if len(first_v[0]) in (2, 3):
                feat["is_paired"] = True

    return feat


# ══════════════════════════════════════════════════════════
# 决策树
# ══════════════════════════════════════════════════════════
def auto_select(data: dict, intent: str | None = None,
                top_k: int = 3) -> list[dict]:
    """根据数据 + 意图返回 top K 推荐。

    Args:
        data: 数据字典
        intent: 表达意图(中英文皆可),可空时根据 data 自动识别
        top_k: 返回前几个候选

    Returns:
        list of {"type": str, "kind": "native_chart"|"diagram",
                 "reason": str, "confidence": 0~1}
    """
    feat = analyze_data(data)
    intent_norm = normalize_intent(intent)

    # 若 intent 为空,根据 feat 自动推断
    if intent_norm is None:
        intent_norm = _infer_intent(feat)

    # 调度到具体决策子树
    if intent_norm == "comparison":
        candidates = _decide_comparison(feat)
    elif intent_norm == "composition":
        candidates = _decide_composition(feat)
    elif intent_norm == "trend":
        candidates = _decide_trend(feat)
    elif intent_norm == "relationship":
        candidates = _decide_relationship(feat)
    elif intent_norm == "distribution":
        candidates = _decide_distribution(feat)
    elif intent_norm == "flow":
        candidates = _decide_flow(feat)
    elif intent_norm == "hierarchy":
        candidates = _decide_hierarchy(feat)
    elif intent_norm == "schedule":
        candidates = _decide_schedule(feat)
    else:
        # 兜底
        candidates = [{
            "type": "column", "kind": "native_chart",
            "reason": "未识别明确意图,默认柱状图",
            "confidence": 0.3,
        }]

    return candidates[:top_k]


def _infer_intent(feat: dict) -> str:
    """根据特征自动推断意图。"""
    if feat["is_flow"]:
        return "flow"
    if feat["is_hierarchy"]:
        return "hierarchy"
    if feat["is_paired"]:
        return "relationship"
    if feat["is_time"]:
        return "trend"
    if feat["value_sums_to_100"]:
        return "composition"
    return "comparison"  # 默认比较


# ══════════════════════════════════════════════════════════
# 各意图的决策子树
# ══════════════════════════════════════════════════════════
def _decide_comparison(feat: dict) -> list[dict]:
    """比较类:柱状/条形/雷达。"""
    n = feat["n_cats"]
    n_series = feat["n_series"]

    if n_series >= 2:
        # 多系列对比
        if n <= 4:
            return [
                _rec("grouped_col", "native_chart",
                     f"{n_series} 个系列 × {n} 类目,分组柱状图最直接", 0.9),
                _rec("stacked_bar", "native_chart",
                     "若想看占比变化,堆叠条形图更合适", 0.7),
                _rec("radar", "native_chart",
                     "若维度独立(评分类),雷达图也是选项", 0.5),
            ]
        else:
            return [
                _rec("grouped_col", "native_chart",
                     f"{n_series} 系列 × {n} 类目,分组柱状图", 0.85),
                _rec("line", "native_chart",
                     "类目较多时折线图更易读", 0.7),
                _rec("stacked_bar", "native_chart",
                     "若关注占比贡献", 0.6),
            ]

    # 单系列
    if n <= 6:
        return [
            _rec("column", "native_chart",
                 f"{n} 个类目对比,垂直柱状图最常用", 0.95),
            _rec("bar", "native_chart",
                 "若类目名较长,水平条形图更易读", 0.7),
            _rec("doughnut", "native_chart",
                 "若数值之和有意义(占比),环形图也行", 0.4),
        ]
    else:
        return [
            _rec("bar", "native_chart",
                 f"{n} 个类目较多,水平条形图避免标签拥挤", 0.9),
            _rec("column", "native_chart",
                 "若标签短,垂直柱也可", 0.6),
            _rec("treemap", "diagram",
                 "若有层次或想强调相对大小,矩形树图", 0.5),
        ]


def _decide_composition(feat: dict) -> list[dict]:
    """占比类:饼/环/堆叠/树图/旭日。"""
    n = feat["n_cats"]
    if n <= 5:
        return [
            _rec("doughnut", "native_chart",
                 f"{n} 个组成部分,环形图清晰",
                 0.9),
            _rec("pie", "native_chart",
                 "饼图更传统但视觉效果接近", 0.85),
            _rec("stacked_bar", "native_chart",
                 "横向 100% 堆叠条更易读数值", 0.7),
        ]
    elif n <= 10:
        return [
            _rec("stacked_bar", "native_chart",
                 f"{n} 类(超过 5),堆叠条形图比饼图易读",
                 0.85),
            _rec("treemap", "diagram",
                 "矩形树图按面积比例直观",
                 0.8),
            _rec("doughnut", "native_chart",
                 "勉强可用,但建议合并小项", 0.5),
        ]
    else:
        return [
            _rec("treemap", "diagram",
                 f"{n} 个类目太多,矩形树图最直观", 0.9),
            _rec("stacked_bar", "native_chart",
                 "若有时间维度,堆叠条/柱可表达变化", 0.6),
        ]


def _decide_trend(feat: dict) -> list[dict]:
    """趋势类:折线/面积/主题河流/K线。"""
    n_series = feat["n_series"]

    if n_series == 1:
        return [
            _rec("line", "native_chart",
                 "单系列时间序列,折线图首选", 0.95),
            _rec("area", "native_chart",
                 "强调累计量时用面积图", 0.7),
            _rec("column", "native_chart",
                 "类目少且想强调离散值,柱状图也可",
                 0.6),
        ]

    if n_series >= 4:
        return [
            _rec("line", "native_chart",
                 f"{n_series} 个系列,折线图避免堆叠混乱", 0.9),
            _rec("theme_river", "native_chart",
                 "若强调总量构成变化,主题河流图(堆叠面积)合适", 0.75),
            _rec("area", "native_chart",
                 "百分比堆叠面积可看占比变化", 0.6),
        ]

    return [
        _rec("line", "native_chart",
             f"{n_series} 系列时间序列,折线图直接", 0.9),
        _rec("area", "native_chart",
             "若强调累计/占比,堆叠面积图", 0.75),
        _rec("theme_river", "native_chart",
             "主题河流图(中心对称)有视觉冲击", 0.5),
    ]


def _decide_relationship(feat: dict) -> list[dict]:
    """关系类:散点/气泡。"""
    sample = None
    if "series" in feat or True:
        # 我们假设 data["series"] 里第一个点的 tuple 长度是 2 或 3
        # 但 feat 里没记,这里从 is_paired 推断
        pass

    return [
        _rec("scatter", "native_chart",
             "二变量相关性,散点图首选", 0.9),
        _rec("bubble", "native_chart",
             "若有第三个变量(大小),气泡图", 0.85),
        _rec("heatmap", "diagram",
             "若两变量都是离散类,热力图", 0.6),
    ]


def _decide_distribution(feat: dict) -> list[dict]:
    """分布类:热力/散点。"""
    return [
        _rec("heatmap", "diagram",
             "二维稠密分布,热力图最直接", 0.9),
        _rec("scatter", "native_chart",
             "若数据点不密集,散点图也行", 0.7),
        _rec("bubble", "native_chart",
             "若想加第三维,气泡图", 0.5),
    ]


def _decide_flow(feat: dict) -> list[dict]:
    """流向类:漏斗/流程图/桑基。"""
    if feat["is_flow"] and "left_nodes" in str(feat):
        # 双层流向 → 桑基
        return [
            _rec("sankey", "diagram",
                 "双层流向数据,桑基图直观", 0.9),
            _rec("stacked_bar", "native_chart",
                 "若简化为左右两栏占比,堆叠条也行", 0.5),
        ]

    # 单向衰减(转化漏斗)
    return [
        _rec("funnel", "diagram",
             "单向衰减(如转化漏斗),漏斗图首选", 0.9),
        _rec("process_flow", "diagram",
             "若强调步骤而非数值,横向流程图", 0.8),
        _rec("bar", "native_chart",
             "若想保留可编辑数据,水平条形图也能模拟", 0.5),
    ]


def _decide_hierarchy(feat: dict) -> list[dict]:
    """层级类:树图/旭日。"""
    return [
        _rec("treemap", "diagram",
             "层级 + 占比,矩形树图最直观",
             0.9),
        _rec("sunburst", "diagram",
             "强调多层放射结构,旭日图视觉感更强",
             0.8),
        _rec("stacked_bar", "native_chart",
             "若只有两层,堆叠条形图也行", 0.5),
    ]


def _decide_schedule(feat: dict) -> list[dict]:
    """排程类:甘特/时间轴。"""
    return [
        _rec("gantt", "diagram",
             "任务排程,甘特图标准", 0.95),
        _rec("process_flow", "diagram",
             "若只是阶段顺序而非具体日期,流程图", 0.6),
    ]


def _rec(chart_type: str, kind: str, reason: str, confidence: float) -> dict:
    """构造推荐项。"""
    return {
        "type": chart_type,
        "kind": kind,
        "reason": reason,
        "confidence": round(confidence, 2),
    }


# ══════════════════════════════════════════════════════════
# 命令行入口(给 Claude / 用户用)
# ══════════════════════════════════════════════════════════
def explain(recommendations: list[dict]) -> str:
    """格式化推荐结果为可读文本(用于 ckpt_3_visual 提示)。"""
    lines = ["📊 推荐图表(按置信度排序):", ""]
    for i, r in enumerate(recommendations, 1):
        kind_icon = "🟢" if r["kind"] == "native_chart" else "🟡"
        lines.append(
            f"  {i}. {kind_icon} {r['type']:15s} [{r['kind']}]  "
            f"({r['confidence']*100:.0f}%)"
        )
        lines.append(f"     理由:{r['reason']}")
        lines.append("")
    lines.append("🟢 = 原生 Chart(可在 PPT 双击改数据)")
    lines.append("🟡 = 形状组合(改要回代码)")
    return "\n".join(lines)


if __name__ == "__main__":
    # 演示
    test_cases = [
        ("时间序列趋势", {
            "cats": ["1月", "2月", "3月", "4月"],
            "series": {"GMV": ([100, 120, 140, 130], None)},
        }),
        ("品牌 SOV 占比", {
            "cats": ["品牌 A", "品牌 B", "品牌 C", "品牌 D"],
            "vals": [38, 23, 24, 15],
        }),
        ("营销漏斗", {
            "stages": ["曝光", "点击", "加购", "下单", "复购"],
            "values": [10000, 3000, 800, 200, 50],
            "is_flow_explicit": True,
        }, "flow"),
    ]

    for case in test_cases:
        if len(case) == 2:
            label, data = case
            intent = None
        else:
            label, data, intent = case
        print(f"\n=== {label} ===")
        recs = auto_select(data, intent=intent)
        print(explain(recs))
