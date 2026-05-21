"""生成本 skill 兼容的最简 PPTX 模板。

生成的 starter.pptx 含:
- 16:9 画布(13.33 × 7.5 in)
- 主题字体已预设为微软雅黑 Light
- 9 个标准 layout(继承 python-pptx 默认 + 字体补丁)
- layout[0] = Title Slide(含标题占位符)
- layout[1] = Title and Content(含标题占位符 idx=0)
- layout[2] = Section Header(章节扉页)
- ...

用法:
    python3 scripts/tools/generate_starter_template.py templates/starter.pptx

如果你已经有真实模板,优先用真实的;starter 仅作为兜底。
"""

import sys
from pathlib import Path

# 把 scripts 加入路径
THIS_DIR = Path(__file__).parent
ROOT = THIS_DIR.parent.parent
sys.path.insert(0, str(ROOT))

from pptx import Presentation
from pptx.util import Inches

from scripts.core.font import patch_theme_fonts


def generate(output_path: str):
    """生成 starter.pptx。"""
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)

    prs = Presentation()
    # 强制 16:9
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # 注入微软雅黑 Light 主题字体
    patch_theme_fonts(prs)

    # 不预填任何 slide(让 assemble 时按需添加)
    prs.save(str(out))

    print(f"✅ Starter template 已生成: {out}")
    print(f"📐 画布: 13.33 × 7.5 in (16:9)")
    print(f"🔤 主题字体: 微软雅黑 Light")
    print(f"📑 Layouts: {len(prs.slide_layouts)} 个(python-pptx 默认布局 + 字体补丁)")
    print()
    print("接下来:")
    print(f"  1. 校验兼容性: python3 scripts/tools/validate_template.py {out}")
    print(f"  2. 在 assemble 中使用: assemble(reg, template_path='{out}', ...)")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 generate_starter_template.py <output_path.pptx>")
        sys.exit(1)
    generate(sys.argv[1])
