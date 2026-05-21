# 13 — Image Provider(B-Visual 路径)v0.5.0

## 核心理念:生图必须可切换 + 比例精确 + 风格专业

本 skill 把"配图生成"抽象为 `ImageProvider` 接口,所有调用方对接同一接口,
实现可热替换。**v0.5.0 新增 aspect 系统 + 专业风格预设**,解决:
- ❌ 旧:1024×1024 图被强行拉伸到 16:9 slide,变形
- ❌ 旧:prompt 太通用("blue-purple gradient"),AI 软糯感

## 四种 Provider

```
ImageProvider (Protocol)
    ├── ManualImageProvider     ← 默认,手动产图后回填(零依赖)
    ├── LocalAssetProvider      ← 本地素材库匹配
    ├── OpenAIImageProvider     ← ⭐ 真实 API 生图
    └── MCPImageProvider        ← Phase 2 占位(独立仓库)
```

### 统一接口(v0.5.0)

所有 provider 的 `generate` 接受 `aspect` 参数,内部处理:

```python
img_path = provider.generate(
    prompt="...",
    aspect="16:9",       # 业务比例
    quality="medium",    # low / medium / high
)
# 返回:精确比例的 PNG 路径(无需 slide 端拉伸)
```

**aspect 参数支持**:
| aspect | API size(实际请求) | 是否后处理裁切 |
|---|---|---|
| `1:1` | 1024×1024 | 否 |
| `2:3` | 1024×1536 | 否(直出) |
| `3:2` | 1536×1024 | 否(直出) |
| **`16:9`** | 1536×1024 | **裁顶/底 → 1536×864** |
| `9:16` | 1024×1536 | 裁左/右 |
| `4:3` / `3:4` | 1536×1024 / 1024×1536 | 裁后输出 |
| `21:9` | 1536×1024 | 裁后输出超宽 |

## 配置(安全)

⚠️ **API key 严禁写入代码** —— 用 `.env` 或环境变量。

### 方式 1:`.env` 文件(推荐)
```bash
cp .env.example .env
# 编辑 .env,填入真实值
```

`.env`:
```
PPTX_CJK_IMAGE_API_BASE=https://your-relay.com/v1
PPTX_CJK_IMAGE_API_KEY=sk-...
PPTX_CJK_IMAGE_MODEL=gpt-image-2
```

`.env` 已在 `.gitignore` 排除,不会被 git 提交。

### 方式 2:环境变量(临时)
```bash
PPTX_CJK_IMAGE_API_KEY=sk-... python3 examples/04_with_real_images.py
```

## 风格预设系统(v0.5.0 — 专业设计参考)

旧版 prompt 过于通用,AI 出图软糯。新版引用具体设计语言:

| 风格 key | 设计参考 | 推荐 aspect | 场景 |
|---|---|---|---|
| `hero_editorial` | Wallpaper / Bloomberg Businessweek 杂志封面 | 16:9 | 提案封面 |
| `hero_keynote_dark` | Apple Keynote 2024 黑底单光源 | 16:9 | 新品发布 |
| `hero_swiss_grid` | Pentagram / Müller-Brockmann 瑞士国际风 | 16:9 | 设计师视角 |
| `hero_brutalist_bold` | Awwwards 2024 新粗野主义 | 16:9 | 前卫商务 |
| `hero_chinese_modern` | 上下 / 高端中国奢侈品(水墨 + 留白) | 16:9 | 中国市场 |
| `concept_growth_metric` | Bloomberg 信息图 3D 上升弧线 | 16:9 | 业绩增长 |
| `concept_journey_path` | NYT explainer + Apple Maps 等距 3D | 16:9 | 客户旅程 |
| `concept_transformation` | Frame magazine 静物摄影 | 16:9 | 转型隐喻 |
| `concept_funnel_abstract` | D3 + ZBrush 数据艺术 | 16:9 | 营销漏斗 |
| `audience_lifestyle_cn` | T Magazine 编辑摄影(中国都市) | 3:2 | 受众洞察 |
| `audience_persona_flat` | New Yorker + Tom Froese 编辑插画 | 1:1 | 人物画像 |
| `product_studio_premium` | Apple 官网产品图 + Hasselblad | 16:9 | 产品 hero |
| `background_color_field` | Mark Rothko 色场抽象 | 16:9 | 内容页底色 |
| `background_paper_texture` | Aesop / Kinfolk 杂志风 | 16:9 | 触感纸质底 |
| `abstract_data_sculpture` | Refik Anadol / Memo Akten 数据艺术 | 16:9 | 数据流可视化 |

### 用法

```python
from scripts.pipeline.image_styles import build_prompt
from scripts.pipeline.image_provider import get_provider

# v0.5.0:返回 (prompt, aspect) 元组
prompt, aspect = build_prompt(
    subject="AcmeCorp 2026 H2 品牌战略",
    style="hero_editorial",
    accent_hex="#7B61FF",       # ⭐ hex 强约束,优于模糊词
    industry="consumer tech",
)

provider = get_provider("openai")
img_path = provider.generate(prompt, aspect=aspect, quality="medium")
# → 精确 16:9 比例,直接铺满 16:9 slide 不变形
```

### Prompt 工程改进点

| 维度 | 旧(v0.4.0) | 新(v0.5.0) |
|---|---|---|
| 设计参考 | 形容词("modern, clean") | 具体作品("Wallpaper magazine", "Pentagram poster") |
| 颜色 | "blue-purple gradient" | hex `#7B61FF` |
| 构图 | "centered composition" | "60% negative space on right, single focal element on left third" |
| 反向约束 | 无 | "Avoid: AI-generic gradients, soft pastel mush" |
| subject 处理 | 直接拼到 prompt | 作为概念引导(避免中文字符画到图里) |

## 集成到 Asset Registry

```python
from scripts.core.registry import attach_visual, KIND_RASTER_PNG

prompt, aspect = build_prompt(...)
img_path = provider.generate(prompt, aspect=aspect)

attach_visual(
    reg, "slide_01",
    kind=KIND_RASTER_PNG,
    path=str(img_path),
    prompt=prompt,           # 复盘用
    style="hero_editorial",  # 风格 key
    aspect=aspect,           # ⭐ 比例信息
    editable_via="mcp_image_edit",
)
```

### Registry 中 raster_png 的标准字段

```json
{
  "kind": "raster_png",
  "path": "work/images/gen_001.png",
  "prompt": "Editorial design hero image, in the style of...",
  "style": "hero_editorial",
  "aspect": "16:9",
  "editable_via": "mcp_image_edit"
}
```

`aspect` 字段让 assemble 阶段知道应该如何放图(铺满 vs aspect-fit)。

## 在三管道工作流里的位置

```
B 视觉管道(B-Visual 子分支)
    │
    ├─ B3 image_provider.generate() ← 用 aspect 控制比例
    │     ├─ ManualImageProvider(零成本)
    │     ├─ LocalAssetProvider(本地素材)
    │     └─ OpenAIImageProvider(真生图)
    │
    ├─ B3.5 PIL 后处理裁切到精确比例(自动)
    │
    ├─ B4 attach_visual(reg, ..., kind=KIND_RASTER_PNG, aspect=...)
    │
    └─ ckpt_3_visual 用户审查后进入 C3 组装
```

C3 组装时 `_build_section_divider_with_hero` 检测图实际比例:
- 若与 slide ratio 差 < 2%,直接铺满(完美匹配)
- 若差 > 2%,走 `add_picture_safe` aspect-fit(留白不变形)

## 实战工作流

### Phase 1(当前 v0.5.0)
```python
# 1. 配置 .env 填入 API key
# 2. 在 example 里:

from scripts.pipeline.image_provider import get_provider
from scripts.pipeline.image_styles import build_prompt
from scripts.core.registry import attach_visual, KIND_RASTER_PNG

provider = get_provider("openai")

for slide_id, subject, style in slides:
    prompt, aspect = build_prompt(
        subject=subject, style=style,
        accent_hex="#1E6BB8",
    )
    img_path = provider.generate(prompt, aspect=aspect)
    attach_visual(reg, slide_id, kind=KIND_RASTER_PNG,
                  path=str(img_path), prompt=prompt,
                  style=style, aspect=aspect)
```

完整示例见 [examples/04_with_real_images.py](../examples/04_with_real_images.py)。

### Phase 2(mcp-image-gen 落地后)
```python
provider = get_provider("mcp")
# 其他代码完全不变
```

## 常见问题

### Q: 生图很慢
- `gpt-image-2` 1536×1024 medium 质量约 60-180 秒
- 优化:用 `quality="low"` 加速到 30-60 秒(本 skill 默认值)
- 真实交付项目用 `medium` 或 `high`

### Q: 中文字符出现在生成的图里
**原因**:旧版直接把中文 subject 拼到 prompt 末尾。
**v0.5.0 修复**:`build_prompt` 把 subject 作为"概念引导"传给模型,加入约束
"convey the mood, not the literal words" + "no text overlays"。

### Q: sandbox / 受限网络环境
本实现自动适配:
- 检测 HTTPS_PROXY 含 `localhost` 时自动改 `127.0.0.1`(Python socket 解析问题)
- 用 `requests` 而非 `openai` SDK,避免 httpx 的 SOCKS 依赖

### Q: 模型选择
| 模型 | 特点 | 适用 |
|---|---|---|
| `gpt-image-1` / `gpt-image-2` | 文字渲染最准、image_edit 强、跟 prompt 对话能力强 | 大部分场景首选 |
| `gemini-3-pro-image-preview` | Google 模型,某些场景独特风格 | 备选 |
| `dall-e-3` | 老牌稳定 | 兼容性好 |
| `qwen-image-2.0-pro` | 国产强 prompt 跟随,中文场景友好 | 中文密集场景 |

### Q: 多个 deck 项目共用配置
每个项目自己的 `.env`,互不干扰。或在 shell 里 export 全局变量。

### Q: 怎么追踪每张图复盘
Registry 已经存了 prompt + style + aspect,改稿时直接基于这些做 image_edit。
还可以从 `gen_001.prompt.txt`(Manual mode 占位文件)里查 prompt。

## 关于其他视觉资产

### Canva / SVG 图标
不走 image_provider —— Canva 没公开 generate API,图标资产通常预选好。
建议:
1. 在 `assets/icons/` 维护图标库
2. registry 用 `kind=KIND_SVG_ICON` + `path=...` 引用
3. assemble 时由 `components.image.add_picture_safe` 处理

### 截图(产品 UI / 平台界面)
直接放 `assets/screenshots/`,registry 用 `kind=KIND_IMAGE_EXTRACTED`。

## 测试覆盖

```bash
python3 tests/test_image_provider.py
# → 19 passed (Protocol / aspect / 4 个 provider / styles / build_prompt 等)
```

## 文件清单

| 文件 | 作用 |
|---|---|
| `scripts/pipeline/image_provider.py` | 4 种 provider + ASPECT_TO_API_SIZE + crop_to_aspect |
| `scripts/pipeline/image_styles.py` | 15+ 专业设计风格预设 + build_prompt(返回 tuple) |
| `tests/test_image_provider.py` | 19 个 mock/单元测试 |
| `.env.example` | 配置模板 |
| `examples/04_with_real_images.py` | 端到端示例(4 张真生图) |
