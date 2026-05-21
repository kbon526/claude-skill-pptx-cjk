# 14 — 模板要求(强约束)

## 强制要求

**本 skill 自 v0.3.0 起强制要求 `template_path` 必填**,不再支持空白创建。
原因:
- python-pptx 默认空白 Presentation 的 layout 占位符约定混乱,渲染效果不可控
- 中文商务 PPT 几乎都基于既有模板(品牌 VI / 客户既有模板 / 行业标准模板)
- 强制模板可以保证 deck 的画布尺寸、母版字体、品牌 chrome 一致

## 模板必须满足的条件

| 条件 | 必需 | 检查方式 |
|---|---|---|
| 16:9 画布(13.33" × 7.5") | ⭐⭐⭐ 强烈推荐 | `validate_template.py` 自动检查 |
| 至少 2 个 layout | ✅ 必需 | 自动检查 |
| 至少 1 个 layout 含 Blank 类型 | ✅ 必需 | skill 用 layout_idx=6 (Blank) |
| 主题字体非锁定 | ⭐⭐ 推荐 | skill 会自动 patch 为微软雅黑 Light |

**实际上要求很低** —— skill 用 Blank layout + 自己画所有元素,不依赖模板的占位符。
模板的价值在于:
1. **画布尺寸**(16:9 vs 4:3)
2. **母版字体**(已通过 `patch_theme_fonts` 注入微软雅黑 Light)
3. **色板 / 主题颜色**(影响图表默认色调)
4. **顶部/底部装饰条 chrome**(slide master 自带的品牌元素)

## 三种获取模板的方式

### 方式 1:用本仓库的 starter 模板

最快路径,适合刚上手:

```bash
# 一键生成兼容模板
python3 scripts/tools/generate_starter_template.py templates/starter.pptx

# 校验(应该看到全 ✅)
python3 scripts/tools/validate_template.py templates/starter.pptx
```

### 方式 2:用客户既有模板

实际客户场景最常见。使用前先校验:

```bash
python3 scripts/tools/validate_template.py /path/to/client_template.pptx
```

输出示例:
```
📂 模板: client_template.pptx
📐 画布: 13.33in × 7.50in
   ✅ 16:9 比例(本 skill 标准)
📑 Layouts: 8 个
   ✅ layout[0] '封面' placeholders=4 title=['idx=0']
   ✅ layout[1] '内容页' placeholders=3 title=['idx=10']
   ...
✅ 模板完全兼容,可直接使用
```

### 方式 3:从客户 PPT 反推模板

如果客户给了一份既有 PPT 但没单独给模板:

```bash
# 用 inspect_template.py 解包查看
python3 scripts/inspect_template.py /path/to/client.pptx

# 把 client.pptx 直接当模板用
# (skill 在 assemble 时会清空 client.pptx 的既有 slides,只保留 layouts)
```

## 模板兼容性常见问题

### Q: 模板是 4:3 比例,可以用吗
**A**: 可以,但本 skill 的网格常数(SAFE_L=0.5、FULL_W=12.33 等)按 16:9 设计,4:3 模板会出现:
- 内容超出右边界(因为 4:3 画布宽度只有 10in,而 FULL_W=12.33in)
- 建议改用 16:9 模板或修改 `scripts/core/layout.py` 常数

### Q: 模板的 layout 6 不是 Blank
**A**: 通常 PowerPoint 默认 9-11 个 layouts,layout 6 是 Blank。但定制模板可能不同。
- 解决:用 `inspect_template.py` 查看哪个 layout 是空白 / 最干净的
- 在调用 `add_content_slide` 时显式传 `layout_idx=<你的 blank idx>`

### Q: 模板有顶部/底部装饰条(chrome),要不要去掉
**A**: 不需要去掉。skill 会保留 master 上的 chrome,只清空 slide 实例上的 placeholders。
- 装饰条作为母版元素会自动出现在每张 slide 上

### Q: 模板的色板想保留,但字体想改成微软雅黑 Light
**A**: 这正是 skill 的默认行为。`assemble` 在加载模板后会:
1. 调用 `patch_theme_fonts(prs)` 把主题字体强制改为微软雅黑 Light
2. 保留所有非字体的色板/主题设置
3. 通过 `_sf()` 注入器保证每个文本 run 的字体一致

## starter.pptx 详情

仓库自带 `templates/starter.pptx`,通过 `generate_starter_template.py` 生成:

- 画布:13.33 × 7.5 in (16:9)
- 主题字体:微软雅黑 Light(全部 latin/eastAsia/cs)
- 11 个标准 layout(继承 python-pptx 默认布局 + 字体补丁):
  - 0: Title Slide
  - 1: Title and Content
  - 2: Section Header
  - 3: Two Content
  - 4: Comparison
  - 5: Title Only
  - **6: Blank** ⭐ skill 默认用这个
  - 7: Content with Caption
  - 8: Picture with Caption
  - 9: Title and Vertical Text
  - 10: Vertical Title and Text

starter 是**功能性模板**,无任何品牌 chrome / 色板 / 装饰元素,纯白底。
适合用作"代码完全控制视觉"的场景。

## 真实生产模板的额外考虑

如果你把本 skill 用于客户提案场景,真实模板通常应该有:

### 必备元素
- **封面 layout**:含品牌 logo 占位、大标题、副标
- **章节扉页 layout**:有章节编号、大标题、副标
- **内容页 layout**:顶部色条、底部页脚 + 页码、品牌色 accent
- **附录页 layout**:简化样式,放数据来源/方法论

### 推荐元素
- 三色配色(主色 + 强调色 + 中性色)预定义为主题色
- 品牌 logo 在 master 上,自动出现在每页
- 页码占位符

### 避免的元素
- 大幅装饰图(挤占内容区)
- 复杂渐变背景(会影响图表可读性)
- 锁定主题字体(干扰本 skill 的字体注入)

## 调试技巧

### 模板能加载但渲染有问题
```bash
# 用 inspect_template.py 查看每个 layout 的 placeholders
python3 scripts/inspect_template.py templates/your.pptx

# 查看 master 元素
python3 -c "from pptx import Presentation; prs = Presentation('templates/your.pptx'); \
  print([s.name for s in prs.slide_master.shapes])"
```

### 字体在 PowerPoint 显示对,但 Keynote 显示错
检查模板是否有 `<a:latin typeface='+mj-lt'/>` 这种"主题字体引用"。
本 skill 的 `_sf()` 直接写死字体名,不会受影响,但模板 master 元素可能受影响。

### 客户模板加载报错
```python
from pptx import Presentation
try:
    prs = Presentation("client.pptx")
except Exception as e:
    print(e)
```
常见原因:
- 文件损坏(用 PowerPoint 打开后另存为可修复)
- 包含特殊形状(Smart Art / 复杂图表),用 `inspect_template.py` 查看具体哪页有问题
