# 07 — slideLayout XML 修复

某些字体 / 字号 / 占位符问题不在 build 脚本层能修，必须直接编辑 `slideLayout*.xml`。这篇收录最常见的两类 XML 修复。

## 修复 1：标题占位符高度不足

### 症状

中文长标题（>14 字）写入 placeholder 后被压扁、换行错位、或直接截断。在 PowerPoint Windows 端尤其严重。

### 根因

模板原作者按英文短标题设计 `slideLayout.xml` 中 title 的 `<a:ext cy="..."/>`，cy 值往往只有 452438 EMU（约 1.26cm），中文长标题塞不下。

### 解法

直接修改 `slideLayoutN.xml`：

```python
import zipfile
import shutil
from lxml import etree

def fix_title_height(pptx_path, layout_index=2, new_cy_emu=648000):
    """
    修复 slideLayoutN.xml 的 title placeholder 高度。
    new_cy_emu = 648000 (约 1.80cm) 适合 24pt 中文双行标题
    """
    # PPTX 是 ZIP 包，直接重写
    layout_path = f"ppt/slideLayouts/slideLayout{layout_index}.xml"
    
    with zipfile.ZipFile(pptx_path, 'r') as zin:
        with zin.open(layout_path) as f:
            tree = etree.parse(f)
    
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
          "p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
    
    # 找到 title 占位符（type="title"）
    title_sp = tree.find(".//p:sp[p:nvSpPr/p:nvPr/p:ph[@type='title']]", ns)
    if title_sp is None:
        # 也可能 type="ctrTitle"
        title_sp = tree.find(".//p:sp[p:nvSpPr/p:nvPr/p:ph[@type='ctrTitle']]", ns)
    
    ext = title_sp.find(".//a:ext", ns)
    ext.set("cy", str(new_cy_emu))
    
    # 写回 ZIP
    new_xml = etree.tostring(tree, xml_declaration=True,
                              encoding="UTF-8", standalone=True)
    
    # 用临时文件 + 替换的方式重写 zip
    tmp_path = pptx_path + ".tmp"
    with zipfile.ZipFile(pptx_path, 'r') as zin:
        with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                if item == layout_path:
                    zout.writestr(item, new_xml)
                else:
                    zout.writestr(item, zin.read(item))
    shutil.move(tmp_path, pptx_path)
```

### 推荐 cy 值

| 标题字号 | 单行 cy | 双行 cy |
|---------|---------|---------|
| 24pt | 432000 (1.20cm) | 648000 (1.80cm) |
| 28pt | 504000 (1.40cm) | 756000 (2.10cm) |
| 32pt | 576000 (1.60cm) | 864000 (2.40cm) |
| 36pt | 648000 (1.80cm) | 972000 (2.70cm) |

EMU 换算：1 cm = 360000 EMU，1 英寸 = 914400 EMU。

## 修复 2：默认字号未定义（继承不稳定）

### 症状

不同 slide 的标题字号不一致。明明在脚本里都用了同一个 layout，结果有的 24pt 有的 28pt。

### 根因

`slideLayout.xml` 中 `defRPr`（default run properties）缺少 `sz` 属性：

```xml
<!-- ❌ 没指定字号 -->
<a:defRPr/>

<!-- ✅ 显式指定 -->
<a:defRPr sz="2400"/>
```

注意 `sz` 单位是百分点（centi-point），24pt = sz="2400"。

没指定时字号继承自 master 或 theme，不同 PowerPoint 版本可能解析出不同值。

### 解法

```python
def fix_title_default_size(pptx_path, layout_index=2, sz_centipoint=2400):
    layout_path = f"ppt/slideLayouts/slideLayout{layout_index}.xml"
    # 同上读 XML
    ...
    
    # 找到 title 占位符的 defRPr
    title_sp = tree.find(".//p:sp[p:nvSpPr/p:nvPr/p:ph[@type='title']]", ns)
    def_rpr = title_sp.find(".//a:lstStyle/a:lvl1pPr/a:defRPr", ns)
    if def_rpr is None:
        # 创建一个
        ...
    def_rpr.set("sz", str(sz_centipoint))
    # 写回
```

## 修复 3：占位符 type 误设

### 症状

明明是 title 占位符，但写入后位置错位 / 字号不对。

### 根因

`<p:ph type="title"/>` vs `<p:ph type="ctrTitle"/>` vs `<p:ph idx="0"/>`（无 type）三者样式不同：
- `title` = 普通标题（左对齐）
- `ctrTitle` = 居中标题（封面用）
- 无 type 仅有 idx = 通用占位符（继承默认样式）

### 解法

明确指定 type：

```xml
<!-- 内容页标题 -->
<p:ph type="title" sz="quarter"/>

<!-- 封面标题 -->
<p:ph type="ctrTitle"/>
```

## 修复 4：装饰元素挡住内容

### 症状

新加的 textbox 被某个母版上的大色块挡住。

### 根因

母版（slideMaster1.xml）或 layout 上有装饰图形（背景色块、Logo 大背景），它们的 z-order 在 layout 内容之上。

### 解法

**方案 A**：把那个装饰元素从 layout 复制到一个独立 layout（仅封面 / 节标题用），其他 layout 不带它。

**方案 B**：在 build 脚本里把新 textbox 显式提到最前：

```python
from pptx.oxml.ns import qn

def bring_to_front(shape):
    sp_tree = shape._element.getparent()
    sp_tree.remove(shape._element)
    sp_tree.append(shape._element)
```

**方案 C**：直接在 layout XML 删掉那个装饰元素的 `<p:sp>` 节点（最暴力）。

## 修复 5：默认字体丢失（lang 属性问题）

### 症状

中文文字在 PowerPoint 里显示成系统默认字体（宋体），而不是模板设定的微软雅黑。

### 根因

`<a:rPr lang="en-US">` 设错了 lang，导致渲染引擎走了西文字体路径。

### 解法

确保中文文本的 run 属性用 `zh-CN`：

```xml
<a:r>
  <a:rPr lang="zh-CN" altLang="en-US" sz="2400">
    <a:latin typeface="Calibri"/>
    <a:ea typeface="微软雅黑"/>
  </a:rPr>
  <a:t>中文标题</a:t>
</a:r>
```

或在 python-pptx 层面：

```python
from pptx.oxml.ns import qn
rPr = run._r.get_or_add_rPr()
rPr.set("lang", "zh-CN")
rPr.set("altLang", "en-US")
```

## XML 编辑工作流

1. **先备份**：`cp template.pptx template.pptx.bak`
2. **解包**：`unzip template.pptx -d template_unpack/`
3. **修改**：用 VS Code / nvim 直接编辑 `ppt/slideLayouts/slideLayoutN.xml`
4. **打包回去**：
   ```bash
   cd template_unpack && zip -r ../template.pptx . && cd ..
   ```
   注意：必须在 `template_unpack/` 目录里 `zip -r ../`，不能 `zip -r template.pptx template_unpack/`（多一层目录就废了）
5. **验证**：在 PowerPoint / Keynote 打开，看是否报"文件已损坏"

## 反例

❌ **直接用 PowerPoint GUI 改 layout**：保存后 PowerPoint 会大量重写 XML，破坏其他你已设好的属性
❌ **修改 master 而不是 layout**：master 改动会影响所有 layouts，连带破坏其他版式
❌ **忘记备份**：XML 改坏整个 .pptx 就废了

## 速查：常用 EMU 换算

```
1 英寸 = 914400 EMU
1 cm   = 360000 EMU
1 mm   = 36000 EMU
1 pt   = 12700 EMU (字号 24pt = sz="2400" 是 centi-point，不是 EMU)
```

---

上一篇：[06-qa.md](06-qa.md) | 下一篇：[08-shape-rebuild.md](08-shape-rebuild.md)
