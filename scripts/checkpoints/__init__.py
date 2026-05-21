"""强制 Checkpoint 系统 —— 三管道工作流的 4 个人工审查门禁。

为什么需要 checkpoint:
1. PPT 制作链路长(框架 → 内容 → 视觉 → 组装),错误越早发现成本越低
2. AI 生成的内容可能"看起来对"但偏离客户需求,必须人工把关
3. 大改稿的痛苦远高于多轮小确认(用户 BambuLab V4→V5 16 处标题改写经验)

四个 checkpoint 节点:
- ckpt_1_framework: A1 框架确认(deck 大纲、章节结构、核心论点)
- ckpt_2_content: A3 内容确认(每页 title/copy/数据)
- ckpt_3_visual: B4 视觉确认(图表选型/AI 生图/素材清单)
- ckpt_4_assembly: C3 组装确认(.pptx 草稿,QA 通过)

⚠️ 重要:Claude 在主对话里推进时,必须"真的停下"等用户回复。
checkpoint 函数会打印清晰的产出物 + 明确询问 → Claude 在收到用户明确批准
("继续"/"OK"/"go" 等)前,不应进入下一阶段。
"""

from . import _gate
from . import ckpt_1_framework
from . import ckpt_2_content
from . import ckpt_3_visual
from . import ckpt_4_assembly

__all__ = [
    "_gate",
    "ckpt_1_framework",
    "ckpt_2_content",
    "ckpt_3_visual",
    "ckpt_4_assembly",
]
