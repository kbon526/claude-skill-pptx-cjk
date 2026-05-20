# 03 — python-pptx 关键坑位与 Workaround

这里收录在多个真实项目中**反复踩过**的坑。看完这篇能少走 3 天弯路。

## 坑 1：macOS 默认 python3 不含 python-pptx

**症状**：
```bash
$ python3 build.py
ModuleNotFoundError: No module named 'pptx'
```

**根因**：macOS 默认的 `python3` 在 `/Library/Developer/CommandLineTools/usr/bin/python3`，是 Python 3.9，site-packages 里没有 `pptx`。

**解法**：显式使用 Python 3.13（或你本机装了 python-pptx 的版本）：

```bash
which python3
# /Library/Developer/CommandLineTools/usr/bin/python3   ← 这个不行

ls /Library/Frameworks/Python.framework/Versions/
# 3.13   ← 用这个

PY=/Library/Frameworks/Python.framework/Versions/3.13/bin/python3
$PY -c "import pptx; print(pptx.__version__)"
# 1.0.2   ← 可用
$PY build.py
```

**建议**：在 build 脚本顶部直接写好 shebang：

```python
#!/Library/Frameworks/Python.framework/Versions/3.13/bin/python3
```

或者用 `uv` / `pyenv` 管理（更优雅）。

## 坑 2：脚本文件名不能与标准库冲突

**症状**：
```
ImportError: cannot import name 'Presentation' from partially initialized module 'pptx'
(most likely due to a circular import)
```

**根因**：你把脚本命名为 `inspect.py`，与 Python 内置 `inspect` 模块同名。lxml 内部 `import inspect` 时优先加载了你的本地脚本，引发循环导入。

**禁用名单**（不完全列举）：
- `inspect.py`、`types.py`、`io.py`、`copy.py`、`time.py`、`string.py`
- `xml.py`、`json.py`、`csv.py`、`html.py`
- `pptx.py`（直接覆盖了你想用的库！）

**解法**：用业务名重命名：
- `inspect_template.py`、`pptx_inspect.py`
- `analyze_pptx.py`、`build_deck.py`

## 坑 3：`add_slide()` 只能末尾追加

**症状**：你想在第 5 页之后插入新页，但 `prs.slides.add_slide(layout)` 永远把新 slide 加到最后。

**解法**：手动操作 `_sldIdLst` XML：

```python
def insert_slide_at(prs, idx, layout):
    """在指定索引位置插入新 slide。"""
    new_slide = prs.slides.add_slide(layout)  # 先追加到末尾
    sld_lst = prs.slides._sldIdLst
    new_sld_id = sld_lst[-1]  # 取出末尾的新 slide ID
    sld_lst.remove(new_sld_id)
    sld_lst.insert(idx, new_sld_id)
    return new_slide
```

完整脚本见 [`scripts/insert_slide_at.py`](../scripts/insert_slide_at.py)。

**注意**：这是直接操作 lxml element，未来 python-pptx 升级可能改 API。我们用的是 1.0.x 验证过。

## 坑 4：删除 slide 同样没有 API

**症状**：找不到 `prs.slides.remove(slide)` 这种方法。

**解法**：

```python
def remove_slide(prs, idx):
    sld_lst = prs.slides._sldIdLst
    sld_id_elem = sld_lst[idx]
    rId = sld_id_elem.get(qn("r:id"))
    prs.part.drop_rel(rId)
    sld_lst.remove(sld_id_elem)
```

`qn` 来自 `pptx.oxml.ns.qn`。

## 坑 5：`shape.text = "..."` 会丢失字体格式

**症状**：你写 `shape.text = "标题"`，结果字体变成默认（宋体或 Calibri），原模板字体丢失。

**根因**：`shape.text = ` 是简便 setter，会清空 text frame 内所有 paragraph 的 run 格式。

**解法**：用更细粒度的 set_text 封装：

```python
def set_text(shape, text, font_name="微软雅黑", size_pt=14,
             bold=False, color=None, align="left"):
    tf = shape.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
                   "right": PP_ALIGN.RIGHT}[align]
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
```

完整版见 [`scripts/set_text.py`](../scripts/set_text.py)。

## 坑 6：海象运算符（`:=`）误用

**症状**：脚本运行时报 `SyntaxError` 或行为诡异。

**真实案例**：在 build_ttd.py 中曾出现 `prs.slide_invariant_or := prs.slide_layouts[1]` 这种误用。

**避免方法**：除非需要在表达式中赋值，否则用普通 `=`。Lint 工具（`ruff`、`pylint`）能捕获大部分误用。

## 坑 7：直接修改 placeholder 后保存丢失某些样式

**症状**：把模板 slide 的 placeholder.text 改了，保存后再打开发现字号变小或字体回退。

**解法**：在改字之前先把字体属性显式 setting 一遍，不要依赖继承：

```python
ph.text = ""  # 先清空
tf = ph.text_frame
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
run = p.add_run()
run.text = "新标题"
run.font.name = "微软雅黑"
run.font.size = Pt(28)
run.font.bold = True
```

## 坑 8：保存后再打开 PowerPoint 报"文件已损坏"

**症状**：脚本生成的 .pptx 在 macOS 上 Keynote 能打开，但 Windows PowerPoint 报错。

**常见原因**：
1. 直接操作 XML 时漏掉某些必需的命名空间
2. 复制 slide 时 rId 关系断了
3. 文件名含特殊字符（中文括号、emoji）

**排查**：
```bash
# 把生成的 pptx 解压，看 [Content_Types].xml 是否完整
unzip -l output.pptx
unzip -p output.pptx "[Content_Types].xml" | head -30
```

或用 `scripts/qa_deep.py` 跑一遍结构检查（详见 [06-qa.md](06-qa.md)）。

## 坑 9：表格行高自动撑大

**症状**：你设了表格 cell 的高度，结果文字一多它自己撑大了。

**解法**：python-pptx 不支持 cell 的 fixed-height，但可以设 word_wrap = False，或者**控制单元格内容长度**避免溢出。

## 坑 10：`pip install` 在公司代理下失败

**症状**：
```
ERROR: Could not find a version that satisfies the requirement python-pptx
ERROR: No matching distribution found for python-pptx
```

**解法**：
1. 先 `pip config list` 看代理设置
2. 试 `pip install --index-url https://pypi.tuna.tsinghua.edu.cn/simple python-pptx`
3. 或者用 `uv pip install python-pptx`，它有自己的镜像逻辑
4. 实在不行，去 PyPI 下载 wheel 离线装：`pip install python_pptx-1.0.2-py3-none-any.whl`

## 坑 11：`markitdown` 对某些 PPTX 静默失败

**症状**：`python -m markitdown some.pptx` 没报错但输出为空。

**解法**：换用 python-pptx 直接读：

```python
from pptx import Presentation
prs = Presentation("some.pptx")
for i, slide in enumerate(prs.slides):
    print(f"--- Slide {i+1} ---")
    for shape in slide.shapes:
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                print(para.text)
```

## 速查：python-pptx API 命名陷阱

| 你以为的 | 实际是 |
|---------|--------|
| `slide.shapes.add_text` | `slide.shapes.add_textbox` |
| `shape.fill.color = RGB(...)` | `shape.fill.solid(); shape.fill.fore_color.rgb = RGBColor(...)` |
| `font.color = "FF0000"` | `font.color.rgb = RGBColor.from_string("FF0000")` |
| `Inches(1.5)` 是 cm | 是英寸（PPTX 内部 EMU = 914400/英寸）|

---

上一篇：[02-template-clone.md](02-template-clone.md) | 下一篇：[04-cjk-fonts.md](04-cjk-fonts.md)
