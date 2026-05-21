"""A3 内容分配 — markdown 切片到 slide_map。

输入:统一 markdown(由 content_ingest 产出)
输出:slide_map.json,每个 slide 含 title / copy / hint(布局建议)

切片规则:
- # 一级标题 → 章节扉页(divider)
- ## 二级标题 → 内容页标题
- 段落正文 → bullet 或 block
- 数字 + 单位组合 → 自动识别为 KPI 候选

⚠️ 这一步只做"草稿映射",最终 slide_map 必须经过 Checkpoint 2(内容确认)
人工审查后才能进入视觉管道。

使用方式:
    from scripts.pipeline.content_allocate import allocate
    slide_map = allocate(markdown_text)
    # → 写入 work/slide_map.json,等用户确认
"""

import re
from pathlib import Path
from typing import Iterator

from ..core.registry import init_registry, add_slide


# 常见 KPI 数字模式(粗筛,不追求精准)
KPI_PATTERNS = [
    re.compile(r"\d+(?:\.\d+)?[%‰]"),                     # 23.5%
    re.compile(r"\d+(?:\.\d+)?\s?[xX×倍]"),                # 5x / 5倍
    re.compile(r"\d+(?:,\d{3})+"),                          # 1,234,567
    re.compile(r"[¥$€£]\s?\d+(?:\.\d+)?[万亿KMB]?"),       # ¥123万
    re.compile(r"\d+(?:\.\d+)?\s?[万亿KMB]"),              # 1.5亿
]


def is_kpi_candidate(text: str) -> bool:
    """判断一行文本是否可能是 KPI(短 + 含数字)。"""
    text = text.strip()
    if len(text) > 30 or len(text) < 2:
        return False
    return any(p.search(text) for p in KPI_PATTERNS)


def split_sections(md: str) -> Iterator[tuple[int, str, list[str]]]:
    """按 # / ## 切分章节,返回 (level, title, body_lines)。"""
    lines = md.split("\n")
    cur_level = 0
    cur_title = ""
    cur_body: list[str] = []

    for raw in lines:
        line = raw.rstrip()
        if line.startswith("# "):
            if cur_title or cur_body:
                yield cur_level, cur_title, cur_body
            cur_level = 1
            cur_title = line[2:].strip()
            cur_body = []
        elif line.startswith("## "):
            if cur_title or cur_body:
                yield cur_level, cur_title, cur_body
            cur_level = 2
            cur_title = line[3:].strip()
            cur_body = []
        elif line.startswith("### "):
            if cur_title or cur_body:
                yield cur_level, cur_title, cur_body
            cur_level = 3
            cur_title = line[4:].strip()
            cur_body = []
        elif line.strip():
            cur_body.append(line.strip())

    if cur_title or cur_body:
        yield cur_level, cur_title, cur_body


def allocate(md: str, deck_title: str = "", brand: str = "") -> dict:
    """切片 markdown → registry 草稿。

    Args:
        md: 统一 markdown 内容
        deck_title: deck 标题(为空时用 md 的第一个 # 标题)
        brand: 品牌名(影响色板选择)

    Returns:
        Asset Registry dict(slides 已填充 title/copy,但 visuals 为空)
    """
    sections = list(split_sections(md))
    if not deck_title and sections:
        # 用第一个 # 一级标题作为 deck 标题
        for lv, t, _ in sections:
            if lv == 1:
                deck_title = t
                break
        deck_title = deck_title or "Untitled Deck"

    reg = init_registry(deck_title=deck_title, brand=brand)

    slide_idx = 0
    for level, title, body in sections:
        slide_idx += 1
        sid = f"slide_{slide_idx:02d}"

        if level == 1:
            # 章节扉页
            sub = body[0] if body else ""
            add_slide(
                reg, sid,
                title=title,
                subtitle=sub,
                layout="section_divider",
                copy={"main": title, "sub": sub},
            )
            continue

        # 内容页:把 body 拆成 bullet / kpi
        kpis = []
        bullets = []
        paragraphs = []
        for line in body:
            if line.startswith("- ") or line.startswith("* "):
                bullets.append(line[2:].strip())
            elif is_kpi_candidate(line):
                kpis.append(line)
            else:
                paragraphs.append(line)

        layout = "default"
        if kpis and bullets:
            layout = "two_column_kpi_bullets"
        elif kpis:
            layout = "kpi_grid"
        elif bullets:
            layout = "bullets"
        elif paragraphs:
            layout = "paragraph"

        add_slide(
            reg, sid,
            title=title,
            layout=layout,
            copy={
                "main": "\n".join(paragraphs),
                "bullets": bullets,
                "kpis": kpis,
            },
        )

    return reg


if __name__ == "__main__":
    import json
    import sys
    if len(sys.argv) < 2:
        print("Usage: python content_allocate.py <markdown_file> [deck_title] [brand]")
        sys.exit(1)
    md = Path(sys.argv[1]).read_text(encoding="utf-8")
    deck_title = sys.argv[2] if len(sys.argv) > 2 else ""
    brand = sys.argv[3] if len(sys.argv) > 3 else ""
    reg = allocate(md, deck_title, brand)
    print(json.dumps(reg, ensure_ascii=False, indent=2))
