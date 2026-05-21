"""Checkpoint 1 — 框架确认(A1 之后)。

时机:Claude 与用户共创了 deck 大纲后,准备进入素材摄入阶段前。
产出物:work/framework.md,含:
- deck 标题 + 受众 + 目标
- 章节结构(每章核心论点)
- 关键数据点清单
- 视觉风格倾向

作用:避免 Claude 闷头摄入素材后才发现框架方向跑偏。
"""

from pathlib import Path

from . import _gate


def announce(framework_path: str | Path,
              deck_title: str,
              sections: list[str],
              audience: str = "",
              goal: str = ""):
    """触发 framework checkpoint。

    Args:
        framework_path: framework.md 草稿路径
        deck_title: deck 标题
        sections: 章节标题列表
        audience: 受众描述
        goal: deck 目标
    """
    summary_lines = [
        f"Deck: {deck_title}",
    ]
    if audience:
        summary_lines.append(f"受众: {audience}")
    if goal:
        summary_lines.append(f"目标: {goal}")
    summary_lines.append("")
    summary_lines.append("章节结构:")
    for i, s in enumerate(sections, 1):
        summary_lines.append(f"  {i}. {s}")

    return _gate.require_user_confirm(
        ckpt_name="ckpt_1_framework",
        artifact_path=framework_path,
        summary="\n".join(summary_lines),
        next_step="A2 素材摄入 → A3 内容映射(将素材按章节分配到 slide)",
    )


if __name__ == "__main__":
    # 演示
    announce(
        framework_path="work/framework.md",
        deck_title="AcmeCorp 2026 H2 媒介策略",
        sections=["市场背景", "竞品分析", "策略制定", "执行计划", "预算分配"],
        audience="客户营销总监 + CMO",
        goal="敲定 H2 投放方案,获得 800 万预算批复",
    )
