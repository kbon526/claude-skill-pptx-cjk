"""品牌色板冒烟测试。"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from pptx.dml.color import RGBColor

from scripts.core.color import (
    BRAND_COLORS, register_brand, get_brand_palette,
    RED, BLUE, GREEN, WHITE, BLACK,
    CLR_FACEBOOK, CLR_TIKTOK, CLR_XHS,
)


def test_neutral_colors_exist():
    """中性色应该都是 RGBColor 实例。"""
    for c in (RED, BLUE, GREEN, WHITE, BLACK):
        assert isinstance(c, RGBColor)


def test_channel_colors_exist():
    """渠道色覆盖核心平台。"""
    assert isinstance(CLR_FACEBOOK, RGBColor)
    assert isinstance(CLR_TIKTOK, RGBColor)
    assert isinstance(CLR_XHS, RGBColor)


def test_brand_colors_dict_has_chinese_brands():
    """品牌字典必须含中国本土品牌。"""
    for brand in ["BambuLab", "小米", "比亚迪", "Anker"]:
        assert brand in BRAND_COLORS, f"{brand} 缺失"
        assert "primary" in BRAND_COLORS[brand]
        assert isinstance(BRAND_COLORS[brand]["primary"], RGBColor)


def test_register_brand_adds_to_dict():
    """register_brand 应该把品牌加入字典。"""
    register_brand(
        "TestBrand",
        primary=RGBColor(0x12, 0x34, 0x56),
        accent=RGBColor(0xAB, 0xCD, 0xEF),
    )
    assert "TestBrand" in BRAND_COLORS
    assert BRAND_COLORS["TestBrand"]["primary"] == RGBColor(0x12, 0x34, 0x56)
    assert BRAND_COLORS["TestBrand"]["accent"] == RGBColor(0xAB, 0xCD, 0xEF)


def test_get_brand_palette_exact_match():
    """精确匹配品牌名应返回对应色板。"""
    pal = get_brand_palette("BambuLab")
    assert pal["primary"] == BRAND_COLORS["BambuLab"]["primary"]


def test_get_brand_palette_fuzzy_match():
    """模糊匹配:大小写不敏感,允许子串。"""
    pal = get_brand_palette("bambulab")  # 小写
    assert pal["primary"] == BRAND_COLORS["BambuLab"]["primary"]


def test_get_brand_palette_fallback():
    """未知品牌返回默认色板,至少含 primary。"""
    pal = get_brand_palette("XXX_NOT_EXIST")
    assert "primary" in pal


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_") and callable(fn):
            try:
                fn()
                print(f"✅ {name}")
            except Exception as e:
                print(f"❌ {name}: {e}")
