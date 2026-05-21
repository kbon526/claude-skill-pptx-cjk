"""RGB 颜色常量池 + 品牌色字典。

设计原则:
1. **配色哈希化** - 所有颜色都用语义化常量名定义,严禁在调用方写裸 RGB
2. **品牌色与中性色分离** - 品牌色(BRAND_COLORS)按品牌索引,中性色(灰阶)
   全局共享
3. **国内品牌优先** - 默认字典覆盖中文社区高频品牌(BambuLab、比亚迪、小米、
   Anker、Realme 等),其他品牌通过 `register_brand()` 自助扩展

沉淀自 generate_ppt.py(thai-auto-media-model 项目)。

新增项目时的标准做法:
    from scripts.core.color import BRAND_COLORS, register_brand, RGBColor

    # 添加新品牌
    register_brand("Acme", primary=RGBColor(0xFF, 0x00, 0x00))

    # 取用
    palette = get_brand_palette("Acme")
    primary = palette["primary"]
"""

from pptx.dml.color import RGBColor


# ══════════════════════════════════════════════════════════
# 中性色 / 灰阶(全局通用)
# ══════════════════════════════════════════════════════════
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x00, 0x00, 0x00)
INK = RGBColor(0x1A, 0x1A, 0x1A)         # 深墨色,正文标题用
DARK_GRAY = RGBColor(0x33, 0x33, 0x33)   # 正文
TEXT_GRAY = RGBColor(0x66, 0x66, 0x66)   # 次要文本
MED_GRAY = RGBColor(0xBF, 0xBF, 0xBF)    # 边框
LIGHT_GRAY = RGBColor(0xF3, 0xF3, 0xF3)  # 浅色背景
SRC_GRAY = RGBColor(0x99, 0x99, 0x99)    # 来源注释
GRID_GRAY = RGBColor(0xEE, 0xEE, 0xEE)   # 图表网格线
RULE_GRAY = RGBColor(0xD9, 0xD9, 0xD9)   # 分割竖线


# ══════════════════════════════════════════════════════════
# 警示色(全局)
# ══════════════════════════════════════════════════════════
RED = RGBColor(0xE6, 0x00, 0x00)         # 主红(警示/重点)
RED_DARK = RGBColor(0xB8, 0x00, 0x00)    # 深红(hover/active)
GREEN = RGBColor(0x00, 0xA0, 0x50)       # 增长/正向
ORANGE = RGBColor(0xFF, 0x8C, 0x00)      # 警告
BLUE = RGBColor(0x1A, 0x56, 0xDB)        # 信息


# ══════════════════════════════════════════════════════════
# 渠道色(社媒/平台官方色,做 channel mix 图表时用)
# ══════════════════════════════════════════════════════════
CLR_FACEBOOK = RGBColor(0x1B, 0x74, 0xE4)
CLR_INSTAGRAM = RGBColor(0xE1, 0x30, 0x6C)
CLR_TIKTOK = RGBColor(0x25, 0x25, 0x25)
CLR_YOUTUBE = RGBColor(0xFF, 0x00, 0x00)
CLR_LINE = RGBColor(0x06, 0xC7, 0x55)
CLR_WECHAT = RGBColor(0x07, 0xC1, 0x60)
CLR_WEIBO = RGBColor(0xE6, 0x16, 0x2D)
CLR_XHS = RGBColor(0xFF, 0x24, 0x42)     # 小红书
CLR_DOUYIN = RGBColor(0x00, 0x00, 0x00)


# ══════════════════════════════════════════════════════════
# 阶段色(roadmap / phase 图常用)
# ══════════════════════════════════════════════════════════
PHASE_AWARENESS = RGBColor(0x1E, 0x6B, 0xB8)  # 立势(蓝)
PHASE_INTEREST = RGBColor(0xE6, 0x8A, 0x00)   # 冲刺(橙)
PHASE_CONVERSION = RGBColor(0xC8, 0x1A, 0x1A)  # 收割(红)


# ══════════════════════════════════════════════════════════
# 品牌色字典 — 国内社区高频品牌
# ══════════════════════════════════════════════════════════
# 每个品牌至少含 primary,可选 accent / dark / light
BRAND_COLORS = {
    # ── 3C / 数码 ──
    "BambuLab": {
        "primary": RGBColor(0x00, 0xAE, 0x42),
        "accent": RGBColor(0xFF, 0x6B, 0x35),
        "dark": RGBColor(0x1F, 0x1F, 0x1F),
    },
    "Anker": {
        "primary": RGBColor(0x00, 0x7C, 0xC2),
        "accent": RGBColor(0x00, 0xC2, 0xB1),
    },
    "小米": {
        "primary": RGBColor(0xFF, 0x67, 0x00),
        "accent": RGBColor(0x00, 0x00, 0x00),
    },
    "Realme": {
        "primary": RGBColor(0xFE, 0xCC, 0x00),
        "accent": RGBColor(0x00, 0x00, 0x00),
    },
    "vivo": {
        "primary": RGBColor(0x41, 0x5F, 0xFF),
    },
    "OPPO": {
        "primary": RGBColor(0x00, 0xA8, 0x52),
    },
    "华为": {
        "primary": RGBColor(0xC7, 0x00, 0x0B),
    },

    # ── 汽车 / 出行 ──
    "比亚迪": {
        "primary": RGBColor(0xE6, 0x00, 0x12),
        "accent": RGBColor(0x00, 0x4E, 0x9B),
    },
    "蔚来": {
        "primary": RGBColor(0x00, 0xBC, 0xD4),
        "dark": RGBColor(0x1A, 0x1A, 0x1A),
    },
    "理想": {
        "primary": RGBColor(0x00, 0x00, 0x00),
        "accent": RGBColor(0xC8, 0xA2, 0x6B),
    },
    "小鹏": {
        "primary": RGBColor(0x8B, 0x2F, 0xC9),
    },
    "GWM": {
        "primary": RGBColor(0xE6, 0x00, 0x00),
    },
    "Toyota": {
        "primary": RGBColor(0x1A, 0x56, 0xDB),
    },
    "Ford": {
        "primary": RGBColor(0x00, 0x66, 0xCC),
    },
    "Denza": {
        "primary": RGBColor(0xD9, 0x40, 0x40),
    },

    # ── 内容 / 平台 ──
    "Moloco": {
        "primary": RGBColor(0xFF, 0x4B, 0x00),
        "dark": RGBColor(0x0F, 0x0F, 0x0F),
    },
    "字节": {
        "primary": RGBColor(0x00, 0x00, 0x00),
        "accent": RGBColor(0xFF, 0x00, 0x4D),
    },
    "腾讯": {
        "primary": RGBColor(0x00, 0xA8, 0xE1),
    },
    "阿里": {
        "primary": RGBColor(0xFF, 0x66, 0x00),
    },
}


def register_brand(name: str, primary: RGBColor, **extras):
    """注册新品牌色。

    Example:
        register_brand("Acme",
                       primary=RGBColor(0xFF, 0x00, 0x00),
                       accent=RGBColor(0x00, 0xFF, 0x00))
    """
    BRAND_COLORS[name] = {"primary": primary, **extras}


def get_brand_palette(brand_name: str) -> dict:
    """获取品牌色板,支持模糊匹配。

    Returns:
        dict: 至少含 primary,可能含 accent / dark / light
        若未找到品牌,返回包含通用蓝色 primary 的回退色板。
    """
    if brand_name in BRAND_COLORS:
        return BRAND_COLORS[brand_name]
    # 模糊匹配(忽略大小写 + 子串)
    lower = brand_name.lower()
    for k, v in BRAND_COLORS.items():
        if lower in k.lower() or k.lower() in lower:
            return v
    # 回退
    return {
        "primary": BLUE,
        "accent": GREEN,
    }
