"""
extract_docx.py — 从 .docx 文件提取段落文本，零依赖（仅用 stdlib）。

适用场景：
- 公司代理网络下 python-docx 装不上
- markitdown 静默失败
- 只需要纯文本而不要样式

用法:
    python3 extract_docx.py input.docx > output.txt
    python3 extract_docx.py input.docx --json > output.json
    python3 extract_docx.py input.docx --markdown > output.md
"""

import argparse
import json
import sys
import xml.etree.ElementTree as ET
import zipfile

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}


def extract_paragraphs(docx_path):
    """返回 [{"text": "...", "style": "Heading1" | "Normal" | ..., "level": int}, ...]"""
    with zipfile.ZipFile(docx_path, "r") as z:
        with z.open("word/document.xml") as f:
            tree = ET.parse(f)

    root = tree.getroot()
    paragraphs = []

    for p in root.iter(f"{{{NS['w']}}}p"):
        # 提取 style
        style = "Normal"
        pPr = p.find(f"{{{NS['w']}}}pPr")
        if pPr is not None:
            pStyle = pPr.find(f"{{{NS['w']}}}pStyle")
            if pStyle is not None:
                style = pStyle.get(f"{{{NS['w']}}}val", "Normal")

        # 提取所有 w:t（文本节点）
        texts = [t.text or "" for t in p.iter(f"{{{NS['w']}}}t")]
        text = "".join(texts).strip()

        if not text and style == "Normal":
            continue  # 跳过空段

        # 推断 heading 级别
        level = 0
        if style.startswith("Heading"):
            try:
                level = int(style.replace("Heading", ""))
            except ValueError:
                level = 0
        elif style in ("Title",):
            level = 0  # 主标题

        paragraphs.append({"text": text, "style": style, "level": level})

    return paragraphs


def to_markdown(paragraphs):
    lines = []
    for p in paragraphs:
        text = p["text"]
        if not text:
            lines.append("")
            continue
        if p["style"] == "Title":
            lines.append(f"# {text}")
        elif p["level"] > 0:
            lines.append(f"{'#' * (p['level'] + 1)} {text}")
        elif "List" in p["style"]:
            lines.append(f"- {text}")
        else:
            lines.append(text)
        lines.append("")  # 段间空行
    return "\n".join(lines)


def to_plain(paragraphs):
    return "\n".join(p["text"] for p in paragraphs if p["text"])


def main():
    parser = argparse.ArgumentParser(
        description="从 .docx 提取段落文本（零依赖）"
    )
    parser.add_argument("docx_path")
    parser.add_argument("--json", action="store_true", help="输出 JSON")
    parser.add_argument("--markdown", action="store_true", help="输出 Markdown")
    args = parser.parse_args()

    try:
        paragraphs = extract_paragraphs(args.docx_path)
    except (KeyError, FileNotFoundError) as e:
        sys.stderr.write(f"错误：{e}\n")
        sys.exit(1)
    except zipfile.BadZipFile:
        sys.stderr.write("错误：文件不是有效的 .docx (ZIP) 包\n")
        sys.exit(1)

    if args.json:
        print(json.dumps(paragraphs, ensure_ascii=False, indent=2))
    elif args.markdown:
        print(to_markdown(paragraphs))
    else:
        print(to_plain(paragraphs))


if __name__ == "__main__":
    main()
