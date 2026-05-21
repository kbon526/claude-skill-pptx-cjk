# 11 — Checkpoints(强制 stop-and-confirm)

## 这是什么

本 skill 的工作流在 4 个关键节点强制停下,等用户明确批准后才能继续。
这 4 个 checkpoint 不是"可选的 best practice",而是 **SKILL.md 里 CRITICAL
明文规定的契约**。

## 4 个 checkpoint

| # | 名字 | 时机 | 产出物 | 用户应该审查什么 |
|---|---|---|---|---|
| 1 | `ckpt_1_framework` | A1 后 | `work/framework.md` | deck 结构、章节划分、核心论点 |
| 2 | `ckpt_2_content` | A3 后 | `work/registry.json`(slides 已填) | 每页 title/copy/数据是否准确 |
| 3 | `ckpt_3_visual` | B4 后 | `work/registry.json`(visuals 已填) | 图表选型/AI 图风格/素材 |
| 4 | `ckpt_4_assembly` | C3 后 | `output/<deck>.pptx` + 缩略图 | 最终视觉效果、字体一致性 |

## 为什么"强制"而不是"建议"

### BambuLab V4→V5 的痛苦经验
2026-05-19 用户做 BambuLab 媒介策略 V5 时,16 处幻灯片标题需要逐一改写。
**根本原因**:V4 阶段没有强制 ckpt_2,Claude 凭"看起来合理"的判断写了一堆
不符合用户语感的标题,最后只能批量返工。

如果当时强制了 ckpt_2,在 16 个标题刚生成时就会停下让用户确认 —— 用户
当场提 3-5 个调整建议比 V5 阶段全量重写省 80% 时间。

### 正向案例:Moloco CTV
同一周做 Moloco CTV 提案时,显式做了"框架确认 → 内容确认 → 视觉确认"
三段式,每段 5-10 分钟同步,最终交付一遍过。

## 工作机制(契约式)

### 技术现实
Python 脚本无法真的"阻塞" Claude 在对话层面前进 —— Claude 是在
对话循环里推进的,Python 进程只能 print 提示。

### 契约设计
1. `checkpoints/ckpt_*.py` 函数 print 一段醒目的 STOP 提示 + 完整产出物摘要
2. SKILL.md 的 **CRITICAL** 章节规定 Claude 必须在主对话中等用户回复
3. 用户明确批准后,Claude 在 registry 里调用 `set_checkpoint(..., "approved")`
4. 进入下一阶段前,可以 `if not all_approved(reg): raise` 做最后一道防线

## Claude 端的执行模板

```python
# === A1: 框架共创 ===
framework_md = co_create_framework_with_user(...)
Path("work/framework.md").write_text(framework_md, encoding="utf-8")

# === Ckpt 1: 框架确认 ===
from scripts.checkpoints import ckpt_1_framework
ckpt_1_framework.announce(
    framework_path="work/framework.md",
    deck_title="...",
    sections=["市场背景", "竞品分析", ...],
    audience="...",
    goal="...",
)
# Claude 此时必须在主对话里向用户提问:
#   "框架草稿如上,是否继续进入素材摄入阶段?"
# 等待用户回复"继续"/"OK"/"go" 等明确批准。
# 如果用户提修改意见:迭代框架,不要跳过 checkpoint。

# === 用户批准后 ===
from scripts.checkpoints._gate import confirm
confirm(reg, "ckpt_1_framework")

# === A2 + A3:摄入和映射 ===
md = ingest_one("brief.docx")
reg = allocate(md, deck_title="...", brand="...")
save(reg, "work/registry.json")

# === Ckpt 2: 内容确认 ===
from scripts.checkpoints import ckpt_2_content
ckpt_2_content.announce("work/registry.json", reg)
# 等用户回复 → confirm(reg, "ckpt_2_content")

# ... 以此类推到 ckpt_3 / ckpt_4
```

## 用户回复模式

### 批准 → 标记 approved → 进入下一阶段
- "继续 / OK / go / 没问题 / approved / 通过"
- "结构没问题,推进吧"

### 拒绝/修改 → 回到当前阶段迭代,不要进入下一阶段
- "第 3 章拆成两章" → 改 framework
- "slide_05 的 KPI 数字不对" → 改 registry,重新 announce ckpt_2

### 模糊回复 → 主动追问,不要替用户决定
- 用户:"嗯"  → Claude:"是确认继续还是要修改某处?"
- 用户:"看起来不错" → Claude:"理解为批准,即将进入 X 阶段,确认?"

## Checkpoint 之间的关系

```
ckpt_1 框架      ckpt_2 内容       ckpt_3 视觉      ckpt_4 组装
   │                 │                 │                │
   ↓ approved        ↓ approved        ↓ approved       ↓ approved
   A2 摄入 ──── A3 映射 ──── B 管道 ──── C3 组装 ──── 交付
                              │ │
                              └─┴── B-Data + B-Visual 并行
```

任何一个 checkpoint 被 reject,流程回到对应管道的开始(不是回到 ckpt_1)。

## 为什么不在 Python 里 raise 强制阻塞

考虑过的方案:`require_user_confirm` 直接 `input()` 等用户输入。
**否决原因**:
1. Claude Code 通常是非交互式调用脚本(stdin 不是终端)
2. 即使能 input(),用户也不应该在终端里输入,而是在 Claude 主对话里说话
3. 真正的"阻塞"应该来自 Claude 的对话纪律,不是 Python 的 IO

所以 checkpoint 是**双层契约**:
- Python 层:打印清晰提示,等待主对话确认
- Claude 对话层:看到提示后必须停下问用户

## 跨会话恢复

如果用户中途离开,过几小时回来继续:
1. 加载 `work/registry.json`
2. 看 `meta.checkpoint_status` 哪些是 approved/pending
3. 从最早的 pending 处继续

```python
reg = load("work/registry.json")
status = reg["meta"]["checkpoint_status"]
if status["ckpt_1_framework"] != "approved":
    # 框架还没确认,从 A1 开始
    ...
elif status["ckpt_2_content"] != "approved":
    # 内容还没确认,从 A2 / A3 开始
    ...
```
