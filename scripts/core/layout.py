"""16:9 网格常数体系。

设计原则:
1. **网格常数化** - 所有位置/尺寸预先定义为常量,严禁在调用方写裸 Inches(0.5)
2. **语义命名** - SAFE_L 比 LEFT_MARGIN 更短,但保留含义
3. **预设区域** - 提供 left_half() / hero_zone() 等高阶函数返回 (l, t, w, h)

幻灯片尺寸(默认 16:9 宽屏):
    13.333" × 7.5"  (1280 × 720 pt @96dpi)

沉淀自 generate_ppt.py(thai-auto-media-model 项目)。
"""

from pptx.util import Inches, Pt


# ══════════════════════════════════════════════════════════
# 幻灯片画布
# ══════════════════════════════════════════════════════════
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)


# ══════════════════════════════════════════════════════════
# 安全边距 / 全宽
# ══════════════════════════════════════════════════════════
SAFE_L = Inches(0.5)              # 左安全边距
SAFE_R = Inches(12.83)            # 右安全位置(不是边距,是 x 坐标)
SAFE_T = Inches(0.42)             # 顶部
SAFE_B = Inches(7.05)             # 底部内容截止 y
FULL_W = Inches(12.33)            # 全宽内容区(SAFE_R - SAFE_L)
GUTTER = Inches(0.2)              # 通用间距


# ══════════════════════════════════════════════════════════
# 标题区(每页通用顶部)
# ══════════════════════════════════════════════════════════
Y_ANCHOR = Inches(0.42)           # 顶部色锚位置
H_ANCHOR = Inches(0.07)
W_ANCHOR = Inches(0.18)
Y_TITLE = Inches(0.54)            # 主标题 y
H_TITLE = Inches(0.55)
Y_SUB = Inches(1.05)              # 副标题 y
H_SUB = Inches(0.34)


# ══════════════════════════════════════════════════════════
# 内容区
# ══════════════════════════════════════════════════════════
Y_CONTENT = Inches(1.50)          # 内容起点 y
Y_SRC = Inches(7.05)              # 数据来源 y
H_SRC = Inches(0.22)
Y_FOOT = Inches(7.27)             # 页脚 y


# ══════════════════════════════════════════════════════════
# 两栏布局 — 文字优先(左大右小)
# ══════════════════════════════════════════════════════════
LW = Inches(7.1)                  # 左栏宽
RX = Inches(7.9)                  # 右栏 x
RW = Inches(4.93)                 # 右栏宽
COL_GAP_X = Inches(7.7)           # 中缝竖线 x


# 两栏布局 — 截图/可视化优先
LW_IMG = Inches(6.8)
RX_IMG = Inches(7.4)
RW_IMG = Inches(5.43)


# ══════════════════════════════════════════════════════════
# KPI 2×2 网格(右栏)
# ══════════════════════════════════════════════════════════
R_KPI_W = Inches(2.35)
R_KPI_H = Inches(0.75)
R_KPI1_X = RX
R_KPI2_X = RX + Inches(2.45)


# ══════════════════════════════════════════════════════════
# 纵向节奏(右栏组件起始 y)
# ══════════════════════════════════════════════════════════
Y_KPI1 = Inches(1.50)
Y_KPI2 = Inches(2.35)
Y_BULLETS = Inches(2.95)
Y_CHART = Inches(4.35)
Y_TABLE = Inches(4.40)


# ══════════════════════════════════════════════════════════
# 区域预设 — 返回 (l, t, w, h) 元组,直接传给 add_textbox/shape
# ══════════════════════════════════════════════════════════
def full_width(t=Y_CONTENT, h=Inches(5.0)):
    """全宽内容区。"""
    return (SAFE_L, t, FULL_W, h)


def left_half(t=Y_CONTENT, h=Inches(5.0)):
    """左半栏(文字优先布局)。"""
    return (SAFE_L, t, LW, h)


def right_half(t=Y_CONTENT, h=Inches(5.0)):
    """右半栏。"""
    return (RX, t, RW, h)


def left_half_img(t=Y_CONTENT, h=Inches(5.0)):
    """左半栏(图优先布局,留更多右栏给文字)。"""
    return (SAFE_L, t, LW_IMG, h)


def right_half_img(t=Y_CONTENT, h=Inches(5.0)):
    """右半栏(配 left_half_img)。"""
    return (RX_IMG, t, RW_IMG, h)


def three_columns(t=Y_CONTENT, h=Inches(5.0), gap=GUTTER):
    """等宽三栏,返回 [(l1,t,w,h), (l2,t,w,h), (l3,t,w,h)]。"""
    col_w = (FULL_W - gap * 2) / 3
    return [
        (SAFE_L + (col_w + gap) * i, t, col_w, h) for i in range(3)
    ]


def four_columns(t=Y_CONTENT, h=Inches(5.0), gap=GUTTER):
    """等宽四栏。"""
    col_w = (FULL_W - gap * 3) / 4
    return [
        (SAFE_L + (col_w + gap) * i, t, col_w, h) for i in range(4)
    ]


def hero_zone():
    """Hero 大字页:整页除安全边距外全部使用。"""
    return (SAFE_L, SAFE_T, FULL_W, SLIDE_H - SAFE_T - Inches(0.5))


def kpi_grid_2x2(top=Y_KPI1, gap_y=Inches(0.85), gap_x=Inches(0.1)):
    """右栏 2×2 KPI 卡片位置,返回 4 个 (l, t, w, h)。"""
    positions = []
    for row in range(2):
        for col in range(2):
            l = R_KPI1_X + (R_KPI_W + gap_x) * col
            t = top + (R_KPI_H + gap_y) * row
            positions.append((l, t, R_KPI_W, R_KPI_H))
    return positions
