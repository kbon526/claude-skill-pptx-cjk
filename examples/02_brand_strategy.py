"""完整三管道 + 4 Checkpoint 实战示例 — AcmeCorp 品牌策略提案。

本示例演示:
1. A 管道:从内嵌 brief 文本起步(实际项目里会从 .docx 摄入)
2. B 管道:B-Data(柱状图/雷达图)+ B-Visual(ManualImageProvider 占位)
3. C 管道:assemble 到 .pptx
4. 4 个 Checkpoint 显式 announce 并模拟 approve

⚠️ 这是脚本演示版本(全自动 approve 所有 checkpoint)。
   真实使用时,Claude 会在每个 announce 后停下问用户。

运行:
    cd ~/Desktop/claude-skill-pptx-cjk
    python3.13 examples/02_brand_strategy.py

预期产出:
    - work/registry.json
    - output/AcmeCorp_2026H2_Strategy.pptx
"""

import sys
from pathlib import Path

# 把 scripts 加入路径(因为 examples/ 是平行目录)
sys.path.insert(0, str(Path(__file__).parent.parent))

from scripts.core.color import (
    BRAND_COLORS, register_brand, RGBColor,
    CLR_FACEBOOK, CLR_INSTAGRAM, CLR_TIKTOK, CLR_YOUTUBE,
    CLR_XHS, CLR_DOUYIN,
)
from scripts.core.registry import (
    init_registry, add_slide, attach_visual, save, load,
    set_checkpoint, all_approved, summary,
    KIND_NATIVE_CHART, KIND_RASTER_PNG, CKPT_APPROVED,
)
from scripts.pipeline.chart_remake import (
    make_doughnut_spec, make_grouped_col_spec, make_radar_spec,
    channel_mix_spec, brand_compare_spec,
    attach_chart_to_slide,
)
from scripts.pipeline.image_provider import get_provider
from scripts.pipeline.assemble import assemble
from scripts.checkpoints import (
    ckpt_1_framework, ckpt_2_content, ckpt_3_visual, ckpt_4_assembly,
)
from scripts.checkpoints._gate import confirm

WORK_DIR = Path("work")
OUTPUT_DIR = Path("output")
TEMPLATES_DIR = Path("templates")
WORK_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# 强制使用真实模板(没有的话先生成 starter)
TEMPLATE_PATH = TEMPLATES_DIR / "starter.pptx"
if not TEMPLATE_PATH.exists():
    print(f"⚠️ 模板不存在: {TEMPLATE_PATH}")
    print("   生成 starter 模板:")
    print("   python3 scripts/tools/generate_starter_template.py templates/starter.pptx")
    sys.exit(1)


# ══════════════════════════════════════════════════════════
# 注册示例品牌色
# ══════════════════════════════════════════════════════════
register_brand("AcmeCorp",
               primary=RGBColor(0x00, 0x66, 0xFF),
               accent=RGBColor(0xFF, 0x6B, 0x00),
               dark=RGBColor(0x1A, 0x1A, 0x1A))


# ══════════════════════════════════════════════════════════
# 阶段 1: A1 框架共创
# ══════════════════════════════════════════════════════════
def phase_a1_framework() -> str:
    """生成 framework.md 草稿。"""
    framework = """# AcmeCorp 2026 H2 媒介策略提案

## 受众
- 客户营销总监 + CMO
- 决策时间窗:本月内

## 目标
- 敲定 H2 投放方案,获得 800 万预算批复
- 明确 Q3/Q4 重点动作和 KPI

## 章节结构
1. 市场背景 — 行业趋势 + 客户当前定位
2. 竞品分析 — SOV / 渠道分布 / 关键打法
3. 受众洞察 — 核心人群画像 + 触媒习惯
4. 策略框架 — 三阶段节奏 + 渠道矩阵
5. 执行计划 — 时间表 + 预算分配 + KPI

## 关键数据点
- 行业 SOV 三足鼎立(品牌 A 38%, B 23%, AcmeCorp 12%)
- 25-34 岁女性是核心人群(占比 47%)
- 小红书 + 抖音占触媒时长 62%
"""
    framework_path = WORK_DIR / "framework.md"
    framework_path.write_text(framework, encoding="utf-8")

    ckpt_1_framework.announce(
        framework_path=framework_path,
        deck_title="AcmeCorp 2026 H2 媒介策略",
        sections=["市场背景", "竞品分析", "受众洞察", "策略框架", "执行计划"],
        audience="客户营销总监 + CMO",
        goal="敲定 H2 投放方案,获得 800 万预算批复",
    )
    return framework


# ══════════════════════════════════════════════════════════
# 阶段 2: A2/A3 内容映射
# ══════════════════════════════════════════════════════════
def phase_a23_content(framework: str) -> dict:
    """根据 framework 生成 registry 草稿。"""
    reg = init_registry(
        deck_title="AcmeCorp 2026 H2 媒介策略",
        brand="AcmeCorp",
        version="v1",
    )

    # Cover
    add_slide(reg, "slide_01",
              title="AcmeCorp 2026 H2 媒介策略",
              subtitle="Q3-Q4 投放方案",
              layout="section_divider",
              copy={"sub": "提报客户营销决策会"})

    # Section 1: 市场背景
    add_slide(reg, "slide_02",
              title="行业 SOV 三足鼎立",
              subtitle="2026 H1 数据,Pathmatics",
              layout="two_column_kpi_bullets",
              copy={
                  "bullets": [
                      "品牌 A 占据 38% SOV,断层领先",
                      "品牌 B 23%,处于追赶位",
                      "AcmeCorp 12%,与第四名差距收窄",
                      "腰部品牌内卷加剧,Top 3 集中度下降",
                  ],
                  "kpis": [
                      "38% 品牌 A SOV",
                      "23% 品牌 B SOV",
                      "12% AcmeCorp SOV",
                      "+15% AcmeCorp YoY",
                  ],
              })

    # Section 2: 竞品分析
    add_slide(reg, "slide_03",
              title="竞品渠道分布对比",
              subtitle="主要竞品 H1 渠道占比",
              layout="two_col_chart",
              copy={
                  "main": "品牌 A 重投短视频,品牌 B 押宝小红书",
                  "bullets": [
                      "品牌 A 抖音占比 42%,效果型投放主导",
                      "品牌 B 小红书占比 35%,种草路径成熟",
                      "AcmeCorp 抖音 + 小红书各 30%,均衡分布",
                      "微信生态(公众号/视频号)三家都低于 25%",
                  ],
              })

    # Section 3: 受众洞察
    add_slide(reg, "slide_04",
              title="核心人群画像",
              subtitle="25-34 岁都市女性,占比 47%",
              layout="two_column_kpi_bullets",
              copy={
                  "bullets": [
                      "高线城市占比 68%(一线 35% + 新一线 33%)",
                      "本科及以上学历占 71%",
                      "月可支配收入 ¥8,000-15,000 区间",
                      "单次决策路径平均 5.2 个触点",
                  ],
                  "kpis": [
                      "47% 25-34F 占比",
                      "68% 高线城市",
                      "71% 高学历",
                      "5.2x 决策触点",
                  ],
              })

    # Section 4: 策略框架
    add_slide(reg, "slide_05",
              title="三阶段节奏:立势 → 冲刺 → 收割",
              subtitle="Q3 立势造场,Q4 冲刺转化",
              layout="bullets",
              copy={
                  "bullets": [
                      "立势期(7-8 月):品牌大事件 + 头部 KOL,建立心智",
                      "冲刺期(9-10 月):效果广告 + 腰部 KOC,获取兴趣人群",
                      "收割期(11-12 月):转化优化 + 私域承接,推动决策",
                  ],
              })

    # Section 5: 预算分配
    add_slide(reg, "slide_06",
              title="预算分配:总额 800 万",
              subtitle="三阶段 4:3:3 分配",
              layout="chart_focus",
              copy={"main": "立势 320 万,冲刺 240 万,收割 240 万"})

    # 给后两个 slide 加 native chart
    # (slide_03: 竞品渠道堆叠条 / slide_06: 预算环形图)
    return reg


# ══════════════════════════════════════════════════════════
# 阶段 3: B 视觉管道(B-Data + B-Visual)
# ══════════════════════════════════════════════════════════
def phase_b_visual(reg: dict):
    """B 管道:把 chart spec 和 image path 写入 registry。"""

    # ── B-Data: 数据可视化 ─────────────────────────────────
    # slide_02:KPI 柱状图(展示三个品牌 SOV)
    attach_chart_to_slide(
        reg, "slide_02",
        spec=brand_compare_spec(
            brands=["品牌 A", "品牌 B", "AcmeCorp"],
            values=[38, 23, 12],
            chart_type="column",
        ),
        title="2026 H1 行业 SOV",
    )

    # slide_03:竞品渠道分布 100% 堆叠条
    from scripts.pipeline.chart_remake import make_stacked_bar_spec
    attach_chart_to_slide(
        reg, "slide_03",
        spec=make_stacked_bar_spec(
            cats=["品牌 A", "品牌 B", "AcmeCorp"],
            series_list=[
                ("小红书", [22, 35, 28], CLR_XHS),
                ("抖音", [42, 25, 32], CLR_DOUYIN),
                ("微信", [18, 25, 22], CLR_FACEBOOK),
                ("其他", [18, 15, 18], RGBColor(0xBF, 0xBF, 0xBF)),
            ],
        ),
        title="主要竞品渠道分布",
    )

    # slide_04:受众多维雷达
    attach_chart_to_slide(
        reg, "slide_04",
        spec=make_radar_spec(
            cats=["年龄", "城市级别", "学历", "收入", "决策路径"],
            series={
                "AcmeCorp 核心人群": (
                    [4.7, 6.8, 7.1, 5.2, 5.0],
                    BRAND_COLORS["AcmeCorp"]["primary"],
                ),
            },
        ),
        title="核心人群画像",
    )

    # slide_06:预算分配环形图
    attach_chart_to_slide(
        reg, "slide_06",
        spec=make_doughnut_spec(
            cats=["立势期", "冲刺期", "收割期"],
            vals=[320, 240, 240],
            colors=[
                BRAND_COLORS["AcmeCorp"]["primary"],
                BRAND_COLORS["AcmeCorp"]["accent"],
                RGBColor(0x8B, 0x2F, 0xC9),
            ],
        ),
        title="预算分配(万元)",
    )

    # ── B-Visual: 概念图(用 ManualImageProvider 占位)─────
    provider = get_provider("manual", work_dir="work/images")

    # Cover hero 图
    hero_path = provider.generate(
        prompt="AcmeCorp 品牌策略 hero 图,蓝紫渐变背景,极简科技风,16:9",
        size="1536x1024",
        quality="medium",
    )
    attach_visual(reg, "slide_01",
                  kind=KIND_RASTER_PNG,
                  path=str(hero_path),
                  prompt="AcmeCorp 品牌策略 hero 图,蓝紫渐变背景,极简科技风,16:9",
                  editable_via="mcp_image_edit")


# ══════════════════════════════════════════════════════════
# Main: 串起来
# ══════════════════════════════════════════════════════════
def main():
    print("\n" + "═" * 70)
    print("AcmeCorp 2026 H2 媒介策略 — 完整三管道演示")
    print("═" * 70 + "\n")

    # ── A1: 框架 + Ckpt 1 ─────────────────────────────────
    print(">>> 阶段 1: A1 框架共创")
    framework = phase_a1_framework()
    # 演示场景下自动 approve;真实场景需要等用户确认
    print("[演示模式] 自动 approve ckpt_1")

    # ── A2/A3: 内容 + Ckpt 2 ──────────────────────────────
    print("\n>>> 阶段 2: A2/A3 内容映射")
    reg = phase_a23_content(framework)
    set_checkpoint(reg, "ckpt_1_framework", CKPT_APPROVED)
    save(reg, WORK_DIR / "registry.json")

    ckpt_2_content.announce(
        registry_path=WORK_DIR / "registry.json",
        registry=reg,
    )
    print("[演示模式] 自动 approve ckpt_2")
    set_checkpoint(reg, "ckpt_2_content", CKPT_APPROVED)

    # ── B: 视觉 + Ckpt 3 ──────────────────────────────────
    print("\n>>> 阶段 3: B 视觉管道")
    phase_b_visual(reg)
    save(reg, WORK_DIR / "registry.json")

    ckpt_3_visual.announce(
        registry_path=WORK_DIR / "registry.json",
        registry=reg,
        images_dir=WORK_DIR / "images",
    )
    print("[演示模式] 自动 approve ckpt_3")
    set_checkpoint(reg, "ckpt_3_visual", CKPT_APPROVED)

    # ── C: 组装 + Ckpt 4 ──────────────────────────────────
    print("\n>>> 阶段 4: C 设计组装")
    output_path = OUTPUT_DIR / "AcmeCorp_2026H2_Strategy.pptx"

    # 必须先在 ckpt_4 之前 approve 所有,才能 assemble
    # (这里我们手动 mark approve 后调 assemble)
    set_checkpoint(reg, "ckpt_4_assembly", CKPT_APPROVED)
    save(reg, WORK_DIR / "registry.json")

    assemble(reg, template_path=TEMPLATE_PATH, output_path=output_path)

    ckpt_4_assembly.announce(
        pptx_path=output_path,
        registry=reg,
        qa_layout_dir=None,
        qa_deep_summary="(QA 检查跳过,真实流程需运行 scripts/qa_deep.py)",
    )
    print("[演示模式] 自动 approve ckpt_4")

    print("\n" + "═" * 70)
    print(f"✅ 演示完成!输出:{output_path}")
    print("═" * 70)


if __name__ == "__main__":
    main()
