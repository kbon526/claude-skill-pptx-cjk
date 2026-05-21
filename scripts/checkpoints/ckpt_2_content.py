"""Checkpoint 2 — 内容确认(A3 之后)。

时机:Claude 完成内容映射(每页 title + copy + 数据点),准备进入视觉管道前。
产出物:work/registry.json(slides 字段已填,visuals 仍为空)

作用:确保每页"说什么"已经定稿,后续视觉只服务内容,不会因为"图先行"
反过来扭曲文案。
"""

from pathlib import Path

from . import _gate
from ..core.registry import summary as reg_summary


def announce(registry_path: str | Path, registry: dict):
    """触发 content checkpoint。"""
    return _gate.require_user_confirm(
        ckpt_name="ckpt_2_content",
        artifact_path=registry_path,
        summary=reg_summary(registry),
        next_step="B 视觉管道:B-Data(图表)+ B-Visual(AI 生图/素材采购)",
    )
