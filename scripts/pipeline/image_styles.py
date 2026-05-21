"""中文商务 PPT 配图风格预设库(v0.5.0 — 引用专业设计参考)。

设计原则(对比 v0.4.0):
1. ❌ 旧:"Minimalist tech-style hero, clean composition" — 太通用,AI 感强
2. ✅ 新:引用具体设计语言("Pentagram poster style", "Frame.io editorial",
       "Apple Keynote 2024 hero")— 让 AI 抓到具体视觉风格

3. ❌ 旧:"blue-purple gradient" — 颜色模糊
4. ✅ 新:"deep indigo (#1A1A2E) to electric purple (#7B61FF)" — hex 强约束

5. ❌ 旧:形容词堆砌("modern, clean, professional, elegant")
6. ✅ 新:具体构图描述("asymmetric grid", "70% negative space",
       "single focal element on left third")

7. ❌ 旧:无 aspect 概念
8. ✅ 新:每个风格自带推荐 aspect ratio + 适用 PPT 场景

核心 API:
    from scripts.pipeline.image_styles import build_prompt, STYLES

    prompt, aspect = build_prompt(
        subject="AcmeCorp 2026 品牌战略",
        style="hero_editorial",
        accent_hex="#7B61FF",
    )
    # → (prompt, "16:9")

    img = provider.generate(prompt, aspect=aspect)
"""

# ══════════════════════════════════════════════════════════
# 风格预设(v0.5.0 — 专业设计参考)
# ══════════════════════════════════════════════════════════
# 每个 entry 含:
# - prompt: 核心 prompt(英文,引用具体设计参考)
# - aspect: 推荐比例
# - description: 中文场景说明
STYLES: dict[str, dict] = {

    # ══ Hero 类(封面/分节扉页)══════════════════════════
    "hero_editorial": {
        "prompt": (
            "Editorial design hero image, in the style of award-winning magazine "
            "covers (Wallpaper, Frame, Bloomberg Businessweek). "
            "Bold asymmetric composition with 60% negative space on the right. "
            "Single sculpted geometric form on the left third — solid, weighty, "
            "with subtle subsurface light. Matte finish, no gloss. "
            "Sophisticated, considered, confident. "
            "Cinematic depth-of-field. Premium editorial photography quality."
        ),
        "aspect": "16:9",
        "description": "编辑设计风 — 像 Wallpaper、Bloomberg Businessweek 杂志大片",
    },

    "hero_keynote_dark": {
        "prompt": (
            "Apple Keynote 2024 keynote hero slide style. "
            "Deep black background (#000000) with single product-like glowing form "
            "at center — sharp edge light, volumetric ray, no actual product. "
            "Dramatic chiaroscuro lighting. "
            "Ultra-minimal: just one focal element + pitch black. "
            "Premium, theatrical, anticipatory mood."
        ),
        "aspect": "16:9",
        "description": "Apple 发布会风 — 黑底单光源,新品发布会感",
    },

    "hero_swiss_grid": {
        "prompt": (
            "Swiss International Style poster design (Müller-Brockmann, "
            "Pentagram). Strict modular grid, generous white space. "
            "Three rectangular blocks of saturated color in 1:2:3 proportion, "
            "arranged on a 12-column grid. Helvetica-era sensibility. "
            "Mathematical, rational, design-school precise. No text, no decoration. "
            "Print-quality vector cleanliness."
        ),
        "aspect": "16:9",
        "description": "瑞士国际风 — Pentagram 海报设计感,极致网格",
    },

    "hero_brutalist_bold": {
        "prompt": (
            "Neo-brutalist editorial design, in the style of awwwards.com hero "
            "sections circa 2024. Massive bold geometric shape filling 50% of frame "
            "in a single saturated color. Stark contrast against off-white "
            "background (#F5F2EE). Raw, confrontational, contemporary. "
            "No gradients, no soft edges — flat solid color blocks only."
        ),
        "aspect": "16:9",
        "description": "新粗野主义 — Awwwards 风格大块对比色,前卫商务",
    },

    "hero_chinese_modern": {
        "prompt": (
            "Modern Chinese contemporary aesthetic, blending traditional ink wash "
            "(水墨) sensibility with minimalist editorial design. "
            "Single inkwash gesture in the upper right, fading into vast "
            "off-white space. One small accent in vermillion (#C8161D). "
            "Quiet confidence, cultural depth, 留白 (negative space) mastery. "
            "Like a high-end Chinese luxury brand campaign (e.g., Shang Xia)."
        ),
        "aspect": "16:9",
        "description": "中式现代 — 水墨留白 + 现代编辑设计,适合中国市场提案",
    },

    # ══ 概念 / 商业隐喻类 ══════════════════════════════
    "concept_growth_metric": {
        "prompt": (
            "Conceptual editorial illustration for a business growth article. "
            "Style: Bloomberg Businessweek infographic. "
            "Single bold ascending arc rendered as a 3D ribbon, in a saturated "
            "single color, against off-white. Cast shadow grounds it. "
            "Suggests upward momentum without being literal. "
            "No bar chart, no real graph — abstract metaphor only."
        ),
        "aspect": "16:9",
        "description": "增长隐喻 — Bloomberg 风信息图,3D 上升弧线",
    },

    "concept_journey_path": {
        "prompt": (
            "Editorial 3D illustration of a winding path through abstract "
            "topographic landscape, top-down isometric perspective. "
            "Path in saturated color contrasts against muted terrain. "
            "Three small marker points along the path, no text labels. "
            "Style: Apple Maps + The New York Times explainer graphics. "
            "Calm, considered, navigable."
        ),
        "aspect": "16:9",
        "description": "客户旅程 — 等距 3D 地形路径,NYT 信息图风",
    },

    "concept_transformation": {
        "prompt": (
            "Conceptual still life: a single sculptural object mid-transformation, "
            "split between two material states (e.g., crystallizing, dissolving, "
            "morphing). Studio lighting, neutral grey backdrop. "
            "Suggests change, evolution, before-and-after. "
            "Style: Frame magazine cover photography, premium product still life."
        ),
        "aspect": "16:9",
        "description": "转型隐喻 — 物体形态变化,杂志静物摄影感",
    },

    "concept_funnel_abstract": {
        "prompt": (
            "Abstract sculptural funnel form rendered in 3D, descending and "
            "narrowing. Gradient color from saturated top to muted bottom. "
            "Single object on neutral backdrop with soft cast shadow. "
            "Style: D3 + ZBrush, premium data art. "
            "No text, no actual chart — sculptural metaphor only."
        ),
        "aspect": "16:9",
        "description": "营销漏斗隐喻 — 3D 雕塑感,数据艺术风",
    },

    # ══ 受众 / 人物类 ════════════════════════════════
    "audience_lifestyle_cn": {
        "prompt": (
            "Editorial lifestyle photography of contemporary Chinese urban "
            "professional, candid moment. Soft natural window light, modern "
            "interior (Shanghai/Beijing aesthetic). Subject from behind or "
            "three-quarter (no face close-up). Muted color palette, "
            "shallow depth of field. Style: T Magazine China, Vogue editorial. "
            "Authentic, aspirational, not cliché."
        ),
        "aspect": "3:2",
        "description": "中国都市人像 — T Magazine 编辑摄影,自然光不正脸",
    },

    "audience_persona_flat": {
        "prompt": (
            "Editorial flat illustration of a single person in profile, "
            "minimal facial features, contemporary clothing in muted palette. "
            "Style: The New Yorker + Tom Froese. "
            "Off-white background, subtle texture. "
            "Sophisticated, narrative, design-led."
        ),
        "aspect": "1:1",
        "description": "人物画像 — New Yorker 编辑插画,简洁不刻板",
    },

    # ══ 产品 / 工业类 ═══════════════════════════════
    "product_studio_premium": {
        "prompt": (
            "Premium product photography hero shot, single object floating in "
            "void. Rim lighting from upper right, deep shadow on lower left. "
            "Pure black or charcoal backdrop with subtle gradient. "
            "Surface details rendered with crisp specular highlights. "
            "Style: Apple product page + Hasselblad campaign. "
            "Reverent, premium, considered."
        ),
        "aspect": "16:9",
        "description": "产品 hero — Apple 官网产品图风格,黑底高级",
    },

    # ══ 装饰 / 背景类 ═════════════════════════════════
    "background_color_field": {
        "prompt": (
            "Color field abstract painting, large flat areas of two complementary "
            "saturated colors meeting in soft horizon line. "
            "Style: Mark Rothko + minimalist editorial. "
            "Quiet, contemplative, premium. Ideal as slide background. "
            "No text, no objects, pure color."
        ),
        "aspect": "16:9",
        "description": "色场背景 — Rothko 现代抽象,适合做内容页底色",
    },

    "background_paper_texture": {
        "prompt": (
            "Premium textured paper background, off-white (#F8F4ED) with subtle "
            "grain. One soft shadow gradient at upper-left corner. "
            "Style: Luxury brand stationery, Aesop catalogue, "
            "Kinfolk magazine layout. "
            "Quiet, tactile, sophisticated. No text."
        ),
        "aspect": "16:9",
        "description": "纸质纹理底 — Aesop / Kinfolk 杂志风,触感细腻",
    },

    # ══ 抽象数据可视化感 ══════════════════════════════
    "abstract_data_sculpture": {
        "prompt": (
            "3D data sculpture as abstract art, in the style of Refik Anadol or "
            "Memo Akten. Flowing particle streams forming a coherent organic "
            "shape, single accent color against deep navy. "
            "Sophisticated, technological without being clichéd. "
            "Premium data art aesthetic."
        ),
        "aspect": "16:9",
        "description": "数据艺术 — Refik Anadol 风,粒子流构成的雕塑",
    },
}


# ══════════════════════════════════════════════════════════
# 通用质量约束(自动追加到所有 prompt)
# ══════════════════════════════════════════════════════════
QUALITY_CONSTRAINTS = (
    "Ultra-high quality, professional, intentional. "
    "No watermarks, no text overlays, no logos, no signatures. "
    "Avoid: AI-generic gradients, soft pastel mush, generic stock photo look, "
    "cluttered composition, low-quality renders."
)


# ══════════════════════════════════════════════════════════
# Prompt 拼装器(返回 prompt + aspect 元组)
# ══════════════════════════════════════════════════════════
def build_prompt(
    subject: str,
    style: str = "hero_editorial",
    accent_hex: str | None = None,
    accent_color: str | None = None,
    industry: str | None = None,
    extra: str | None = None,
    override_aspect: str | None = None,
) -> tuple[str, str]:
    """组装 prompt + 推荐 aspect。

    Args:
        subject: 业务主题(自然语言,不直接拼到 prompt 末尾,作为概念引导)
        style: 风格 key(见 STYLES)
        accent_hex: 强调色 hex,如 "#7B61FF" — 强约束,优于 accent_color
        accent_color: 强调色描述(英文优先),如 "deep indigo"
        industry: 行业语境,如 "automotive luxury" "consumer tech"
        extra: 额外约束(英文)
        override_aspect: 覆盖默认 aspect,如 "1:1"

    Returns:
        (prompt, aspect) 元组,直接传给 provider.generate()
    """
    if style not in STYLES:
        raise ValueError(
            f"Unknown style: {style}. 可用: {sorted(STYLES.keys())}"
        )

    cfg = STYLES[style]
    parts = [cfg["prompt"]]

    # 主题作为概念引导(不直接说出主题文字,避免 AI 把中文写到图里)
    if subject:
        parts.append(f"Conceptual context: about '{subject}' — convey the mood, not the literal words.")

    # 强调色:优先 hex
    if accent_hex:
        parts.append(f"Primary accent color: {accent_hex} (use as the dominant saturated color).")
    elif accent_color:
        parts.append(f"Primary accent color: {accent_color}.")

    if industry:
        parts.append(f"Industry context: {industry}.")
    if extra:
        parts.append(extra)

    parts.append(QUALITY_CONSTRAINTS)

    aspect = override_aspect or cfg["aspect"]
    return " ".join(parts), aspect


def list_styles() -> dict[str, dict]:
    """列出所有风格 + 描述 + aspect。"""
    return {
        k: {"description": v["description"], "aspect": v["aspect"]}
        for k, v in STYLES.items()
    }


def explain() -> str:
    """打印所有可用风格(用于 ckpt_3_visual 提示)。"""
    cats: dict[str, list[tuple[str, dict]]] = {}
    for k, v in STYLES.items():
        prefix = k.split("_")[0]
        cats.setdefault(prefix, []).append((k, v))

    lines = ["📚 配图风格预设(v0.5.0 — 引用专业设计参考)", ""]
    for cat, items in sorted(cats.items()):
        cat_label = {
            "hero": "Hero 封面 / 分节扉页",
            "concept": "概念 / 商业隐喻",
            "audience": "受众 / 人物",
            "product": "产品 / 工业",
            "background": "背景 / 装饰",
            "abstract": "抽象 / 数据艺术",
        }.get(cat, cat)
        lines.append(f"【{cat_label}】")
        for k, v in items:
            lines.append(f"  - {k}  [{v['aspect']}]")
            lines.append(f"    {v['description']}")
        lines.append("")
    lines.append("用法:")
    lines.append("  prompt, aspect = build_prompt(")
    lines.append("      subject='...', style='hero_editorial',")
    lines.append("      accent_hex='#7B61FF',")
    lines.append("  )")
    lines.append("  img = provider.generate(prompt, aspect=aspect)")
    return "\n".join(lines)


if __name__ == "__main__":
    print(explain())
    print()
    print("=" * 70)
    print("示例 prompt:")
    print("=" * 70)
    p, asp = build_prompt(
        subject="AcmeCorp 2026 H2 brand strategy launch",
        style="hero_editorial",
        accent_hex="#7B61FF",
        industry="consumer tech",
    )
    print(f"Aspect: {asp}\n")
    print(p)
