"""通用 Checkpoint Gate 实现 —— 契约式停顿。

技术现实:Python 脚本无法真的"阻塞" Claude 在对话层面前进。
契约方式:
1. checkpoint 函数打印高亮的 STOP 提示 + 完整产出物摘要
2. SKILL.md 的 CRITICAL 段落规定 Claude 必须等用户明确批准
3. 用户批准后,Claude 在 registry 里调用 set_checkpoint(..., "approved")

Claude 端的执行模式:
    # 1. 完成阶段产出物
    save(reg, "work/registry.json")

    # 2. 调用对应 checkpoint 函数(打印产出物摘要)
    require_user_confirm(
        ckpt_name="ckpt_1_framework",
        artifact_path="work/framework.md",
        summary=summary(reg),
    )

    # 3. 在主对话里向用户提问:"框架已起草完毕,是否继续?"
    # 4. 用户明确说 "继续" 后,在 registry 里 mark approved,进入下一阶段
"""

from pathlib import Path


class CheckpointAborted(Exception):
    """用户拒绝某个 checkpoint。"""
    pass


def require_user_confirm(
    ckpt_name: str,
    artifact_path: str | Path,
    summary: str,
    next_step: str = "",
) -> str:
    """打印 checkpoint 强提示。

    Args:
        ckpt_name: checkpoint 名,与 registry 里的字段对应
        artifact_path: 该阶段产出物路径(.md / .json / .pptx)
        summary: 给用户看的摘要内容
        next_step: 下一阶段会做什么(让用户清楚进度)

    Returns:
        提示字符串(同时已 print)。Claude 收到后应在主对话里向用户征求确认。
    """
    sep = "═" * 70
    msg = f"""
{sep}
🛑 CHECKPOINT: {ckpt_name}
{sep}
产出物: {artifact_path}

────────────────────  摘要  ────────────────────
{summary}
─────────────────────────────────────────────────
{f"⏭️ 下一阶段: {next_step}" if next_step else ""}

⚠️ 在用户明确批准前,Claude 不应进入下一阶段。
   等待用户输入"继续"/"OK"/"go" 等确认信号。
   若用户提出修改意见,在当前阶段内迭代,不要跳过本 checkpoint。
{sep}
"""
    print(msg)
    return msg


def confirm(reg: dict, ckpt_name: str):
    """用户明确批准后调用 — 把 checkpoint 标记为 approved。"""
    from ..core.registry import set_checkpoint, CKPT_APPROVED
    set_checkpoint(reg, ckpt_name, CKPT_APPROVED)


def reject(reg: dict, ckpt_name: str):
    """用户要求重做 — 标记 rejected,通常意味着回到上一阶段重跑。"""
    from ..core.registry import set_checkpoint, CKPT_REJECTED
    set_checkpoint(reg, ckpt_name, CKPT_REJECTED)
