"""pptx-cjk 核心基建模块。

提供四大基础能力:
- font: 中文字体注入器 + 字号体系
- color: RGB 常量池 + 品牌色字典
- layout: 16:9 网格常数体系
- registry: Asset Registry JSON 操作

所有模块都从真实生产项目沉淀,可直接 `from scripts.core import xxx` 使用。
"""

from . import font, color, layout, registry

__all__ = ["font", "color", "layout", "registry"]
