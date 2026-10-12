from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

@dataclass
class Decision:
    """
    一个审计节点的输出结果。

    label：写入底稿的结果，例如：
        是 / N/A / Refer to PMS&VMS / N/A-Refer to PMS/VMS / N/A-容器化无需审阅
    requires_evidence：是否需要后续 OCR / 人工检查凭证。
        - 是：通常需要本系统或当前节点凭证
        - Refer to XXX：需要去 XXX 凭证处检查
        - N/A-Refer to XXX：本节点不直接审阅，但仍需引用 XXX 凭证
        - N/A：通常不需要凭证
    refer_to：如果需要引用其他 Sheet / 系统，在这里写引用对象。
    reason：规则来源或解释。
    """

    label: str
    requires_evidence: bool = False
    refer_to: Optional[str] = None
    reason: str = ""

    @classmethod
    def yes(cls, reason: str = "") -> "Decision":
        return cls("是", requires_evidence=True, reason=reason)

    @classmethod
    def na(cls, reason: str = "") -> "Decision":
        return cls("N/A", requires_evidence=False, reason=reason)

    @classmethod
    def na_with_reason(cls, suffix: str, reason: str = "") -> "Decision":
        return cls(f"N/A-{suffix}", requires_evidence=False, reason=reason)

    @classmethod
    def refer(cls, target: str, reason: str = "") -> "Decision":
        return cls(f"Refer to {target}", requires_evidence=True, refer_to=target, reason=reason)

    @classmethod
    def na_refer(cls, target: str, reason: str = "") -> "Decision":
        return cls(f"N/A-Refer to {target}", requires_evidence=True, refer_to=target, reason=reason)

    @classmethod
    def custom(cls, label: str, requires_evidence: bool = True, reason: str = "") -> "Decision":
        return cls(label, requires_evidence=requires_evidence, reason=reason)

    @classmethod
    def unknown(cls, reason: str = "") -> "Decision":
        return cls("待确认", requires_evidence=False, reason=reason)
