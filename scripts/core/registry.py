"""Asset Registry — 三管道汇合的中心化资产注册表。

核心思想:
- A 内容管道 / B 视觉管道并行产出后,所有资产以 slide_id 为 key 汇总到
  registry,C 组装管道再按 slide_id 一对一映射到幻灯片
- 这样三管道完全解耦,任一管道局部重跑都不破坏整体

JSON Schema:
    {
      "meta": {
        "deck_title": "...",
        "brand": "BambuLab",
        "version": "v1",
        "created_at": "2026-05-20T15:00:00",
        "checkpoint_status": {
          "ckpt_1_framework": "approved",
          "ckpt_2_content": "pending",
          "ckpt_3_visual": "pending",
          "ckpt_4_assembly": "pending"
        }
      },
      "slides": {
        "slide_01": {
          "title": "...",
          "subtitle": "...",
          "layout": "two_column_hero",
          "copy": {
            "main": "...",
            "bullets": ["..."],
            "kpis": [{"value": "5x", "label": "..."}]
          },
          "visuals": [
            {"kind": "native_chart", "spec": {"type": "column", "cats": [...], "vals": [...]}},
            {"kind": "raster_png", "path": "assets/hero.png",
             "prompt": "...", "style": "hero_editorial", "aspect": "16:9",
             "editable_via": "mcp_image_edit"},
            {"kind": "svg_icon", "path": "assets/icon.svg"}
          ],
          "notes": "演讲者备注..."
        }
      }
    }

Visual kind 枚举:
- native_chart: python-pptx 原生图表(B-Data 路径,可编辑)
- raster_png: 栅格图(B-Visual 路径,通常来自 AI 生图)
- svg_icon: 矢量图标(B4 asset sourcing)
- table: 表格数据
- image_extracted: 从素材里提取出来的现成图(B1 路径)
"""

import json
from datetime import datetime
from pathlib import Path
from typing import Any


# Visual kind 常量
KIND_NATIVE_CHART = "native_chart"  # python-pptx 原生 Chart(可双击编辑数据)
KIND_DIAGRAM = "diagram"            # 形状组合(数据嵌入代码,改要回代码)
KIND_RASTER_PNG = "raster_png"
KIND_SVG_ICON = "svg_icon"
KIND_TABLE = "table"
KIND_IMAGE_EXTRACTED = "image_extracted"

VALID_KINDS = {
    KIND_NATIVE_CHART,
    KIND_DIAGRAM,
    KIND_RASTER_PNG,
    KIND_SVG_ICON,
    KIND_TABLE,
    KIND_IMAGE_EXTRACTED,
}


# Checkpoint 状态
CKPT_PENDING = "pending"
CKPT_APPROVED = "approved"
CKPT_REJECTED = "rejected"


def init_registry(deck_title: str, brand: str = "", version: str = "v1") -> dict:
    """初始化一个空的 registry。"""
    return {
        "meta": {
            "deck_title": deck_title,
            "brand": brand,
            "version": version,
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "checkpoint_status": {
                "ckpt_1_framework": CKPT_PENDING,
                "ckpt_2_content": CKPT_PENDING,
                "ckpt_3_visual": CKPT_PENDING,
                "ckpt_4_assembly": CKPT_PENDING,
            },
        },
        "slides": {},
    }


def add_slide(
    reg: dict,
    slide_id: str,
    *,
    title: str = "",
    subtitle: str = "",
    layout: str = "default",
    copy: dict | None = None,
    notes: str = "",
):
    """新增一张 slide。slide_id 建议格式 slide_01 / slide_02 ..."""
    reg["slides"][slide_id] = {
        "title": title,
        "subtitle": subtitle,
        "layout": layout,
        "copy": copy or {},
        "visuals": [],
        "notes": notes,
    }


def attach_visual(
    reg: dict,
    slide_id: str,
    kind: str,
    **payload,
):
    """为某张 slide 附加一个 visual 资产。

    Example:
        # 原生图表
        attach_visual(reg, "slide_03",
                      kind=KIND_NATIVE_CHART,
                      spec={"type": "column", "cats": [...], "vals": [...]})

        # AI 生成的栅格图
        attach_visual(reg, "slide_05",
                      kind=KIND_RASTER_PNG,
                      path="assets/hero_05.png",
                      prompt="cyberpunk city at night",
                      editable_via="mcp_image_edit")
    """
    if kind not in VALID_KINDS:
        raise ValueError(
            f"Unknown visual kind: {kind}. Valid: {sorted(VALID_KINDS)}"
        )
    if slide_id not in reg["slides"]:
        raise KeyError(f"Slide {slide_id} not in registry. Call add_slide first.")
    reg["slides"][slide_id]["visuals"].append({"kind": kind, **payload})


def set_checkpoint(reg: dict, ckpt_name: str, status: str):
    """更新 checkpoint 状态。"""
    if ckpt_name not in reg["meta"]["checkpoint_status"]:
        raise KeyError(f"Unknown checkpoint: {ckpt_name}")
    if status not in (CKPT_PENDING, CKPT_APPROVED, CKPT_REJECTED):
        raise ValueError(f"Invalid status: {status}")
    reg["meta"]["checkpoint_status"][ckpt_name] = status


def all_approved(reg: dict) -> bool:
    """是否所有 checkpoint 都已通过(用于 C3 组装前检查)。"""
    return all(
        s == CKPT_APPROVED for s in reg["meta"]["checkpoint_status"].values()
    )


def save(reg: dict, path: str | Path):
    """落盘为 JSON。"""
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(reg, f, ensure_ascii=False, indent=2)


def load(path: str | Path) -> dict:
    """从 JSON 加载。"""
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def validate(reg: dict) -> list[str]:
    """校验 registry 完整性,返回错误列表(空表示通过)。"""
    errors: list[str] = []

    if "meta" not in reg:
        errors.append("missing meta")
    elif "deck_title" not in reg["meta"]:
        errors.append("meta.deck_title is required")

    if "slides" not in reg:
        errors.append("missing slides")
        return errors

    for sid, s in reg["slides"].items():
        if not sid.startswith("slide_"):
            errors.append(f"slide id should start with 'slide_': {sid}")
        if "title" not in s:
            errors.append(f"{sid} missing title")
        for i, v in enumerate(s.get("visuals", [])):
            if "kind" not in v:
                errors.append(f"{sid}.visuals[{i}] missing kind")
            elif v["kind"] not in VALID_KINDS:
                errors.append(
                    f"{sid}.visuals[{i}] unknown kind: {v['kind']}"
                )
            if v.get("kind") == KIND_RASTER_PNG and "path" not in v:
                errors.append(f"{sid}.visuals[{i}] (raster_png) missing path")

    return errors


def summary(reg: dict) -> str:
    """生成可读的人工审查摘要(用于 checkpoint 对话)。"""
    meta = reg["meta"]
    lines = [
        f"Deck: {meta['deck_title']} ({meta.get('brand', 'no brand')}) {meta['version']}",
        f"Slides: {len(reg['slides'])}",
        "",
    ]
    for sid, s in reg["slides"].items():
        n_visuals = len(s.get("visuals", []))
        kinds = ", ".join(v["kind"] for v in s.get("visuals", []))
        lines.append(
            f"  {sid}: {s['title']!r} [layout={s['layout']}, "
            f"visuals={n_visuals}{f' ({kinds})' if kinds else ''}]"
        )
    lines.append("")
    lines.append("Checkpoints:")
    for ck, status in meta["checkpoint_status"].items():
        emoji = {"approved": "✅", "pending": "⏸", "rejected": "❌"}[status]
        lines.append(f"  {emoji} {ck}: {status}")
    return "\n".join(lines)
