"""三管道编排脚本。

A 内容管道 / B 视觉管道 / C 设计管道的实际执行入口。
Claude 在按 SKILL.md 工作流推进时,会显式调用这些脚本作为步骤产出物。

模块组织:
- content_ingest: A2 摄入(PDF/PPT/docx → markdown)
- content_allocate: A3 内容分配(markdown → slide_map.json)
- visual_extract: B1 视觉提取(从素材提取 figures/charts)
- chart_remake: B2 图表重制(注入品牌色)
- image_provider: B3 生图抽象层(Manual / MCP)
- assemble: C3 组装引擎(Asset Registry → .pptx)
"""
