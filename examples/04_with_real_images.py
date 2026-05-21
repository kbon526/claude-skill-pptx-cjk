"""真实生图端到端示例(v0.5.0 — 修复比例失真 + 提升风格质量)。

⚠️ 运行前置:
1. 复制 .env.example 为 .env,填入 API key
2. 或临时通过环境变量传入:
   PPTX_CJK_IMAGE_API_BASE="https://your-relay.com/v1" \\
   PPTX_CJK_IMAGE_API_KEY="sk-..." \\
   PPTX_CJK_IMAGE_MODEL="gpt-image-2" \\
   python3 examples/04_with_real_images.py

输出:
- output/AcmeCorp_with_AI_Hero.pptx
- work/images/gen_*.png(已 PIL 裁切到精确 16:9)
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from scripts.core.registry import (
    init_registry, add_slide, attach_visual, save,
    set_checkpoint, CKPT_APPROVED, KIND_RASTER_PNG,
)
from scripts.pipeline.image_provider import get_provider
from scripts.pipeline.image_styles import build_prompt
from scripts.pipeline.assemble import assemble


OUTPUT = Path("output/AcmeCorp_with_AI_Hero.pptx")
TEMPLATE = Path("templates/starter.pptx")
WORK_DIR = Path("work/images")
OUTPUT.parent.mkdir(parents=True, exist_ok=True)
WORK_DIR.mkdir(parents=True, exist_ok=True)

if not TEMPLATE.exists():
    print("⚠️ Template missing. Run: python3 scripts/tools/generate_starter_template.py templates/starter.pptx")
    sys.exit(1)


# ══════════════════════════════════════════════════════════
# 4 张 Hero 图任务 — 用 v0.5.0 新风格预设
# ══════════════════════════════════════════════════════════
hero_tasks = [
    # (slide_id, title, subtitle, style, accent_hex, industry)
    ("slide_01",
     "AcmeCorp 2026 H2 品牌战略",
     "Q3-Q4 投放方案 · 2026 年 5 月",
     "hero_editorial",         # 编辑设计风
     "#7B61FF",                # 紫
     "consumer tech"),

    ("slide_02",
     "市场背景 · 行业竞争格局",
     "三足鼎立的竞争态势",
     "concept_growth_metric",  # 增长隐喻 - Bloomberg 风
     "#FF6B35",                # 橙
     "consumer tech"),

    ("slide_03",
     "受众洞察 · 25-34 岁都市女性",
     "47% 核心人群画像",
     "audience_lifestyle_cn",  # 中国都市人像 - T Magazine 风
     "#C8A0A0",                # 柔粉
     "lifestyle"),

    ("slide_04",
     "执行计划 · 三阶段节奏",
     "立势 → 冲刺 → 收割",
     "concept_journey_path",   # 客户旅程 - NYT 信息图风
     "#1E6BB8",                # 科技蓝
     "marketing strategy"),
]


# ══════════════════════════════════════════════════════════
# 是否复用现有图(work/images/ 已有 4 张就跳 API)
# ══════════════════════════════════════════════════════════
GENERATE_REAL_IMAGES = not all(
    (WORK_DIR / f"gen_{i:03d}.png").exists() for i in range(1, 5)
)
QUALITY = "low"  # low / medium / high

print("=" * 70)
print("AcmeCorp 品牌策略 deck + AI Hero 图(v0.5.0 风格 + 比例修复)")
print("=" * 70)

provider = None
if GENERATE_REAL_IMAGES:
    try:
        provider = get_provider("openai", work_dir=str(WORK_DIR), timeout=420)
        print(f"✓ Provider 初始化成功")
        print(f"  api_base: {provider.api_base}")
        print(f"  model:    {provider.model}")
    except RuntimeError as e:
        print(f"❌ {e}")
        sys.exit(1)
else:
    print(f"♻️  检测到 {WORK_DIR}/ 已有 4 张图,复用(跳过 API 调用)")


# ══════════════════════════════════════════════════════════
# 准备 registry
# ══════════════════════════════════════════════════════════
reg = init_registry("AcmeCorp 2026 H2 品牌策略",
                     brand="AcmeCorp", version="v1")


for idx, (slide_id, title, subtitle, style, accent_hex, industry) in enumerate(hero_tasks, 1):
    print(f"\n>>> 处理 {slide_id}: {title}")

    add_slide(reg, slide_id,
              title=title, subtitle=subtitle,
              layout="section_divider",
              copy={"main": title, "sub": subtitle})

    # 用 v0.5.0 build_prompt — 返回 (prompt, aspect) 元组
    prompt, aspect = build_prompt(
        subject=title,
        style=style,
        accent_hex=accent_hex,
        industry=industry,
    )
    print(f"   📝 Style: {style}  Aspect: {aspect}  Accent: {accent_hex}")
    print(f"   📝 Prompt: {prompt[:90]}...")

    img_path = None
    if GENERATE_REAL_IMAGES:
        try:
            img_path = provider.generate(
                prompt=prompt, aspect=aspect, quality=QUALITY,
            )
        except Exception as e:
            print(f"   ⚠️ 生图失败: {e}")
    else:
        img_path = WORK_DIR / f"gen_{idx:03d}.png"
        print(f"   ♻️  复用 {img_path}")

    if img_path and img_path.exists():
        attach_visual(reg, slide_id,
                      kind=KIND_RASTER_PNG,
                      path=str(img_path),
                      prompt=prompt, style=style, aspect=aspect,
                      editable_via="mcp_image_edit")


# ══════════════════════════════════════════════════════════
# 组装
# ══════════════════════════════════════════════════════════
for ck in ["ckpt_1_framework", "ckpt_2_content",
           "ckpt_3_visual", "ckpt_4_assembly"]:
    set_checkpoint(reg, ck, CKPT_APPROVED)

save(reg, "work/with_real_images_registry.json")

print("\n" + "=" * 70)
print(f"组装 {len(reg['slides'])} 张 slide → {OUTPUT}")
print("=" * 70)

assemble(reg, template_path=TEMPLATE, output_path=OUTPUT)

print(f"\n✅ 输出: {OUTPUT}")
print(f"\n💡 生成的图片(已 PIL 裁到精确 16:9,不会在 PPT 里被拉伸):")
for f in sorted(WORK_DIR.glob("gen_*.png")):
    try:
        from PIL import Image
        with Image.open(f) as im:
            w, h = im.size
        ratio = w / h
        print(f"   - {f}  {w}×{h}  ratio={ratio:.3f}  ({f.stat().st_size/1024:.1f} KB)")
    except ImportError:
        print(f"   - {f}  ({f.stat().st_size/1024:.1f} KB)")
