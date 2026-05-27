from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Sequence, Tuple

from decisions import Decision
from field_translator import FieldFacts
from text_tools import TextTools

@dataclass
class ReleaseToolScope:
    front: Decision
    back: Decision
    warnings: List[str] = field(default_factory=list)


class ReleaseToolRuleEngine:
    """
    发版工具前台/后台输出映射。

    这里输出的是 WP 字段值，不是简单 True / False。
    """

    # 应用系统自带发版功能存在系统级口径，不能只按工具名判断。
    # 这些值来自当前 WP / Breakdown 已有结果；如与 Lead Sheet 文字冲突，保留 warning。
    SYSTEM_RELEASE_OVERRIDES: Dict[str, Tuple[Decision, Decision, str]] = {
        "EPM": (
            Decision.custom("由应用层管理员配置发版", True, "应用系统自带发版功能；WP 按应用层管理员配置发版处理。"),
            Decision.na("应用系统自带发版功能后台不适用。"),
            "EPM 为系统级特殊规则。",
        ),
        "HJHJDE": (
            Decision.na("应用系统自带发版功能；WP 中 HJH JDE 前台为 N/A。"),
            Decision.na("应用系统自带发版功能后台不适用。"),
            "HJH JDE 为系统级特殊规则。",
        ),
        "JDE": (
            Decision.yes("应用系统自带发版功能；WP 中 JDE 前台为 是。"),
            Decision.na("应用系统自带发版功能后台不适用。"),
            "注意：Lead Sheet 文字中出现“JDE为否”，但 WP 当前结果为“是”，需最终确认。",
        ),
        "PA-FAS144": (
            Decision.yes("应用系统自带发版功能；WP 中 PA-FAS144 前台为 是。"),
            Decision.na("应用系统自带发版功能后台不适用。"),
            "PA-FAS144 为系统级特殊规则。",
        ),
        "PEOPLESOFT": (
            Decision.yes("应用系统自带发版功能；WP 中 PeopleSoft 前台为 是。"),
            Decision.na("应用系统自带发版功能后台不适用。"),
            "PeopleSoft 为系统级特殊规则。",
        ),
        "TESLA": (
            Decision.custom("应用层Echo账号发版", True, "应用系统自带发版功能；WP 按应用层 Echo 账号发版处理。"),
            Decision.na("应用系统自带发版功能后台不适用。"),
            "TESLA 为系统级特殊规则。",
        ),
        "百川": (
            Decision.yes("应用系统自带发版功能；WP 中百川前台为 是。"),
            Decision.na("应用系统自带发版功能后台不适用。"),
            "百川为系统级特殊规则。",
        ),
    }

    def determine(self, facts: FieldFacts) -> ReleaseToolScope:
        warnings: List[str] = []

        if facts.no_backend:
            return ReleaseToolScope(
                Decision.na("Remark 包含无后台。"),
                Decision.na("Remark 包含无后台。"),
            )

        # Utility Tool：除合规支持平台和 PMS&VMS 外，无需检查发版工具前后台。
        if facts.is_utility_tool and not facts.is_compliance_support_platform and not facts.is_pms_vms:
            return ReleaseToolScope(
                Decision.na("Utility Tool（除合规支持平台和 PMS&VMS）无需检查发版工具前台。"),
                Decision.na("Utility Tool（除合规支持平台和 PMS&VMS）无需检查发版工具后台。"),
            )

        # Other：无需检查发版工具前后台。
        if facts.is_other:
            return ReleaseToolScope(
                Decision.na("Other 无需检查发版工具前台。"),
                Decision.na("Other 无需检查发版工具后台。"),
            )

        # PMS&VMS：需检查发版工具前台，无需检查后台。
        if facts.is_pms_vms:
            return ReleaseToolScope(
                Decision.yes("PMS&VMS 需检查发版工具前台。"),
                Decision.na("PMS&VMS 无需检查发版工具后台。"),
            )

        # SOX / 合规支持平台：按 Lead Sheet 发版工具规则。
        if facts.is_sox_system or facts.is_compliance_support_platform:
            return self._determine_sox_or_compliance(facts)

        warnings.append(f"未覆盖的系统类型：Category={facts.category!r}, System={facts.system!r}。")
        return ReleaseToolScope(
            Decision.unknown("未命中发版工具规则。"),
            Decision.unknown("未命中发版工具规则。"),
            warnings,
        )

    def _determine_sox_or_compliance(self, facts: FieldFacts) -> ReleaseToolScope:
        warnings: List[str] = []

        system_key = TextTools.normalize_key(facts.system)
        if system_key in self.SYSTEM_RELEASE_OVERRIDES:
            front, back, note = self.SYSTEM_RELEASE_OVERRIDES[system_key]
            if "JDE为否" in note:
                warnings.append(note)
            return ReleaseToolScope(front, back, warnings)

        if not facts.pms_flow_items or not facts.release_tool_items:
            return ReleaseToolScope(
                Decision.unknown("SOX/合规系统缺少 PMS流程 或 发版工具。"),
                Decision.unknown("SOX/合规系统缺少 PMS流程 或 发版工具。"),
                [f"缺少 PMS流程或发版工具：System={facts.system!r}"],
            )

        front_decisions: List[Decision] = []
        back_decisions: List[Decision] = []

        for pms_flow in facts.pms_flow_items:
            for release_tool in facts.release_tool_items:
                front, back, local_warnings = self._map_one_tool(pms_flow, release_tool, facts)
                front_decisions.append(front)
                back_decisions.append(back)
                warnings.extend(local_warnings)

        front_merged = self._merge_decisions(front_decisions, node="front")
        back_merged = self._merge_decisions(back_decisions, node="back")

        return ReleaseToolScope(front_merged, back_merged, warnings)

    def _map_one_tool(self, pms_flow: str, release_tool: str, facts: FieldFacts) -> Tuple[Decision, Decision, List[str]]:
        warnings: List[str] = []
        flow = TextTools.normalize_text(pms_flow)
        tool = TextTools.normalize_text(release_tool)
        tool_key = TextTools.normalize_key(tool)

        # 码云发版 / 码云
        if flow == "PMS-企业流程" and tool_key in {"码云发版", "码云"}:
            return (
                Decision.yes("PMS-企业流程 + 码云发版：前台为 是。"),
                Decision.refer("码云", "PMS-企业流程 + 码云发版：后台 Refer to 码云。"),
                warnings,
            )

        if flow == "PMS-电商流程" and tool == "Infra Jenkins公共发布Job":
            return (
                Decision.yes("PMS-电商流程 + Infra Jenkins公共发布Job：前台为 是。"),
                Decision.refer("Infra Jenkins", "PMS-电商流程 + Infra Jenkins公共发布Job：后台 Refer to Infra Jenkins。"),
                warnings,
            )

        if flow == "PMS-电商流程" and tool == "PMS/VMS调用NIO":
            front = Decision.na_refer("PMS/VMS", "PMS-电商流程 + PMS/VMS调用NIO：前台 N/A-Refer to PMS/VMS。")

            if facts.is_app_server_containerized or facts.no_app_layer or facts.no_app_layer_user:
                back = Decision.na_refer(
                    "PMS/VMS",
                    "PMS-电商流程 + PMS/VMS调用NIO；应用层容器化/无应用层用户时，后台 N/A-Refer to PMS/VMS。",
                )
            else:
                back = Decision.refer(
                    "Infra Jenkins",
                    "PMS-电商流程 + PMS/VMS调用NIO；应用层未容器化时，后台 Refer to Infra Jenkins。",
                )
            return front, back, warnings

        if flow == "PMS-企业流程" and tool == "PMS/VMS调用NIO":
            return (
                Decision.yes("PMS-企业流程 + PMS/VMS调用NIO：前台为 是。"),
                Decision.refer("PMS/VMS", "PMS-企业流程 + PMS/VMS调用NIO：后台 Refer to PMS/VMS。"),
                warnings,
            )

        if flow == "PMS-企业流程" and tool == "PMS调用PAM授权":
            return (
                Decision.na("PMS-企业流程 + PMS调用PAM授权：前台 N/A。"),
                Decision.refer("PMS/VMS", "PMS-企业流程 + PMS调用PAM授权：后台 Refer to PMS/VMS。"),
                warnings,
            )

        if flow == "PMS-企业流程" and tool == "应用系统自带发版功能":
            warnings.append(
                f"{facts.system!r} 为应用系统自带发版功能，但未配置系统级覆盖规则；请确认前台账号输出。"
            )
            return (
                Decision.yes("应用系统自带发版功能：默认前台为 是，需确认系统特例。"),
                Decision.na("应用系统自带发版功能：后台 N/A。"),
                warnings,
            )

        if flow == "PMS-企业流程" and tool in {"Jenkins（自建）-纳管", "Jenkins（自建）-已纳管"}:
            return (
                Decision.yes("PMS-企业流程 + Jenkins（自建）-纳管/已纳管：前台为 是。"),
                Decision.yes("PMS-企业流程 + Jenkins（自建）-纳管/已纳管：后台为 是。"),
                warnings,
            )

        if flow == "PMS-企业流程" and tool == "Jenkins（自建）-未纳管":
            return (
                Decision.yes("PMS-企业流程 + Jenkins（自建）-未纳管：前台为 是。"),
                Decision.yes("PMS-企业流程 + Jenkins（自建）-未纳管：后台为 是。"),
                warnings,
            )

        warnings.append(f"未配置发版工具映射：PMS流程={flow!r}, 发版工具={tool!r}, System={facts.system!r}。")
        return (
            Decision.unknown("未配置发版工具映射。"),
            Decision.unknown("未配置发版工具映射。"),
            warnings,
        )

    @staticmethod
    def _merge_decisions(decisions: Sequence[Decision], node: str) -> Decision:
        """
        多个发版工具同时出现时的合并规则。

        观察 WP：
        - 有“是”时，前台通常输出“是”。
        - 后台多引用时，PMS/VMS 优先于 Infra Jenkins，Infra Jenkins 优先于 码云。
        - 只有 N/A-Refer 时保留 N/A-Refer。
        - 全部 N/A 则 N/A。
        """
        if not decisions:
            return Decision.unknown("没有可合并的发版工具判断。")

        labels = [d.label for d in decisions]
        reasons = "；".join([d.reason for d in decisions if d.reason])

        if any(label == "待确认" for label in labels):
            return Decision.unknown(reasons)

        special_labels = [
            label for label in labels
            if label not in {"是", "N/A"}
            and not label.startswith("Refer to ")
            and not label.startswith("N/A-Refer to ")
        ]
        if special_labels:
            return Decision.custom(special_labels[0], True, reasons)

        if node == "front":
            if "是" in labels:
                return Decision.yes(reasons)
            for target in ["PMS/VMS", "Infra Jenkins", "码云"]:
                if f"N/A-Refer to {target}" in labels:
                    return Decision.na_refer(target, reasons)
                if f"Refer to {target}" in labels:
                    return Decision.refer(target, reasons)
            return Decision.na(reasons)

        # 后台：优先保留引用对象。
        for target in ["PMS/VMS", "Infra Jenkins", "码云"]:
            if f"Refer to {target}" in labels:
                return Decision.refer(target, reasons)
        for target in ["PMS/VMS", "Infra Jenkins", "码云"]:
            if f"N/A-Refer to {target}" in labels:
                return Decision.na_refer(target, reasons)
        if "是" in labels:
            return Decision.yes(reasons)
        return Decision.na(reasons)
