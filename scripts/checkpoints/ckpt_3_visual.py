"""Checkpoint 3 — 视觉确认(B4 之后)。

时机:B 管道全部完成(图表 spec 写好 + AI 图产出 + 图标采购完毕),
     准备进入 C 设计组装阶段前。
产出物:work/registry.json(visuals 字段已填满)+ work/images/ 目录

作用:让用户在 .pptx 实际生成前,审查每页将放什么图、什么图表。
这一步通过后,C3 组装时基本不会再有"图选错了"的返工。
"""

from pathlib import Path

from . import _gate


def announce(registry_path: str | Path, registry: dict,
              images_dir: str | Path = "work/images"):
    """触发 visual checkpoint。"""
    lines = [f"Deck: {registry['meta']['deck_title']}", ""]
    lines.append("每页视觉资产:")
    for sid in sorted(registry["slides"].keys()):
        s = registry["slides"][sid]
        visuals = s.get("visuals", [])
        if not visuals:
            lines.append(f"  {sid}: {s['title']!r} — (无图)")
        else:
            kinds = ", ".join(v["kind"] for v in visuals)
            lines.append(f"  {sid}: {s['title']!r} — [{kinds}]")
            for v in visuals:
                if v["kind"] == "raster_png":
                    lines.append(f"      📸 {v.get('path', '?')}")
                elif v["kind"] == "native_chart":
                    spec = v.get("spec", {})
                    lines.append(
                        f"      📊 {spec.get('type', '?')} chart "
                        f"({len(spec.get('cats', []))} cats)"
                    )
                elif v["kind"] == "image_extracted":
                    lines.append(f"      🖼  {v.get('path', '?')} (extracted)")
                elif v["kind"] == "svg_icon":
                    lines.append(f"      🎨 {v.get('path', '?')} (icon)")
    lines.append("")
    lines.append(f"图片目录: {images_dir}")

    return _gate.require_user_confirm(
        ckpt_name="ckpt_3_visual",
        artifact_path=registry_path,
        summary="\n".join(lines),
        next_step="C 设计管道:C1 色板 → C2 模板匹配 → C3 组装为 .pptx",
    )
