"""Checkpoint 4 — 组装确认(C3 之后)。

时机:.pptx 已生成,QA(qa_deep + render_layout)已跑通,准备交付前。
产出物:output/<deck>.pptx + qa_layout/preview-*.png(视觉缩略图)

作用:最终交付前的总验证 — 字体一致性、布局无溢出、图表配色对齐。
通过后才把成品文件交给用户(或写入最终目录)。

⚠️ 这是最后一道关 — 此处发现的问题会回到 B 或 C 管道修复,
   不会再回到内容管道。所以前面 ckpt_2 一定要严格,不要把内容问题
   留到这里。
"""

from pathlib import Path

from . import _gate


def announce(pptx_path: str | Path,
              registry: dict,
              qa_layout_dir: str | Path | None = None,
              qa_deep_summary: str = ""):
    """触发 assembly checkpoint。"""
    pptx_path = Path(pptx_path)
    file_size = "?"
    if pptx_path.exists():
        file_size = f"{pptx_path.stat().st_size / 1024 / 1024:.2f} MB"

    lines = [
        f"Deck: {registry['meta']['deck_title']} ({registry['meta'].get('brand', '')})",
        f"输出文件: {pptx_path} ({file_size})",
        f"幻灯片数: {len(registry['slides'])}",
    ]
    if qa_layout_dir:
        lines.append(f"视觉缩略图: {qa_layout_dir}")
    if qa_deep_summary:
        lines.append("")
        lines.append("结构 QA:")
        for ln in qa_deep_summary.split("\n"):
            lines.append(f"  {ln}")

    return _gate.require_user_confirm(
        ckpt_name="ckpt_4_assembly",
        artifact_path=pptx_path,
        summary="\n".join(lines),
        next_step="交付 — 最终 .pptx 可发送给客户/用户。",
    )
