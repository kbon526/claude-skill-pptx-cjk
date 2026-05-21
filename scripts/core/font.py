"""中文字体注入器 + 字号体系。

核心问题:python-pptx 默认只设置 latin 字体,中文(eastAsia)会回退到
PowerPoint 主题字体(常常是等线/宋体/雅黑 Regular),导致字距过宽、
风格不统一。本模块通过 `_sf()` 统一注入 latin/eastAsia/cs 三处字体,
彻底解决中文不一致问题。

使用规则:
1. 所有 Run 写入文本后,必须调用 `_sf(run, ...)` 注入字体
2. 标题字号用 `SZ_TITLE`,正文用 `SZ_BODY`,以此类推(语义化常量)
3. 不使用 Bold 权重(微软雅黑 Light 在中文场景视觉更克制)
4. 视觉层级靠"字号 + 颜色"区分,而非粗细

主题字体补丁:
- `patch_theme_fonts(prs)` 在 Presentation 加载后调用一次,把主题里
  残留的 等线 / 宋体 / DengXian 全部强制改为 微软雅黑 Light。
- 这样所有继承主题的占位符也会自动用统一字体。

沉淀自 generate_ppt.py(thai-auto-media-model 项目)。
"""

from pptx.util import Pt
from pptx.oxml.ns import qn
from lxml import etree

# ── 字体名 ─────────────────────────────────────────────────
FONT_TITLE = "微软雅黑"        # 占位符标题用
FONT_BODY = "微软雅黑"          # 默认正文
FONT_LIGHT = "微软雅黑 Light"  # ⭐ 全局首选(本 skill 默认)


# ── 字号体系(6 级)─────────────────────────────────────────
# 严格分级,不要随意增加新字号 —— 避免视觉散乱。
SZ_HERO = Pt(32)   # Hero 大字页(封面/分节封面)
SZ_TITLE = Pt(22)  # 主标题(每页结论)
SZ_H3 = Pt(13)     # 组标题 / 列标题
SZ_SUB = Pt(11)    # 副标题(数据支撑)
SZ_BODY = Pt(10)   # 正文 / bullet
SZ_CAP = Pt(9)     # 图表标签 / 表格数据
SZ_SRC = Pt(8)     # 来源 / 页脚


# ── 字体注入器 ────────────────────────────────────────────
def _sf(run, size, bold=False, color=None, title=False, family=FONT_LIGHT):
    """统一字体注入 —— 所有文本写入必须经过此函数。

    Args:
        run: pptx.text.text._Run 对象
        size: pptx.util.Pt(...) 字号
        bold: 历史兼容参数,默认 False(本 skill 不用 Bold,靠字号/颜色分层)
        color: pptx.dml.color.RGBColor 颜色,None 表示继承
        title: 历史兼容参数,语义已合并到 size
        family: 字体名,默认 微软雅黑 Light

    关键操作:
        1. 设置 run.font.name(latin)
        2. 直接操作底层 XML,显式设置 a:ea(eastAsia)和 a:cs(complex script)
           —— 这是中文字体不回退的关键
    """
    run.font.size = size
    run.font.bold = False  # 本 skill 不使用 Bold
    run.font.name = family
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        existing = rPr.find(qn(tag))
        if existing is not None:
            rPr.remove(existing)
        el = etree.SubElement(rPr, qn(tag))
        el.set("typeface", family)
    if color:
        run.font.color.rgb = color


def patch_theme_fonts(prs, target_font=FONT_LIGHT):
    """把 Presentation 主题字体全部改为 target_font。

    覆盖范围:
    - majorFont/minorFont 下的 latin/ea/cs
    - <a:font script='Hans'/Hant/Hang/Jpan'> CJK fallback

    这样所有继承主题的占位符、形状默认字体都会是 target_font,
    避免中文出现等线(DengXian)残留,也避免 Regular 权重。

    在 Presentation 加载后立即调用一次即可:
        prs = Presentation(template)
        patch_theme_fonts(prs)
    """
    ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    cjk_scripts = {"Hans", "Hant", "Hang", "Jpan"}
    for part in prs.part.package.iter_parts():
        if "theme" not in part.partname.lower():
            continue
        try:
            root = etree.fromstring(part.blob)
        except Exception:
            continue
        changed = False
        for tag_name in ("majorFont", "minorFont"):
            for ff in root.iter(f"{{{ns}}}{tag_name}"):
                for child in list(ff):
                    local = child.tag.rsplit("}", 1)[-1]
                    if local in ("latin", "ea", "cs"):
                        child.set("typeface", target_font)
                        changed = True
                    elif local == "font":
                        script = child.get("script") or ""
                        if script in cjk_scripts:
                            child.set("typeface", target_font)
                            changed = True
        if changed:
            part._blob = etree.tostring(
                root, xml_declaration=True, standalone=True, encoding="UTF-8"
            )


# ── 便捷封装:写一段格式化文本到现有 textframe ─────────────
def write_run(paragraph, text, size=SZ_BODY, color=None, family=FONT_LIGHT):
    """在已有 paragraph 上追加一个格式化好的 run。

    用法:
        p = textframe.paragraphs[0]
        write_run(p, "标题", size=SZ_TITLE, color=RED)
    """
    r = paragraph.add_run()
    r.text = text
    _sf(r, size, color=color, family=family)
    return r
