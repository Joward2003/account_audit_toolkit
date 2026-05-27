from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

from audit_nodes import AuditNodes
from rules_scope import ScopeResult

@dataclass
class EvidenceTask:
    system: str
    node: str
    expected_label: str
    evidence_type: str
    refer_to: Optional[str] = None
    expected_keywords: List[str] = field(default_factory=list)
    expected_fields: List[str] = field(default_factory=list)
    reason: str = ""


class EvidenceTaskRouter:
    """
    根据 AccountAudit 输出结果生成凭证任务。

    只有 Decision.requires_evidence=True 的节点才生成 OCR/凭证检查任务。
    """

    NODE_TO_EVIDENCE_TYPE: Dict[str, str] = {
        AuditNodes.RELEASE_FRONT: "release_tool_front_account_evidence",
        AuditNodes.RELEASE_BACK: "release_tool_back_account_evidence",
        AuditNodes.APP_ADMIN: "app_admin_account_evidence",
        AuditNodes.APP_SCREENSHOT: "app_account_screenshot_evidence",
        AuditNodes.APP_SERVER: "app_server_account_evidence",
        AuditNodes.DB_SERVER: "db_server_account_evidence",
        AuditNodes.DB_ACCOUNT: "db_account_evidence",
        AuditNodes.DEV_LIST: "developer_list_evidence",
        AuditNodes.DEV_NO_PROD_ABOVE_READ: "developer_prod_permission_evidence",
        AuditNodes.SOD: "front_back_admin_sod_evidence",
    }

    def build_tasks(self, scope_result: ScopeResult) -> List[EvidenceTask]:
        facts = scope_result.facts
        tasks: List[EvidenceTask] = []

        for node, decision in scope_result.decisions.items():
            if not decision.requires_evidence:
                continue

            target = decision.refer_to or facts.system
            tasks.append(
                EvidenceTask(
                    system=facts.system,
                    node=node,
                    expected_label=decision.label,
                    evidence_type=self.NODE_TO_EVIDENCE_TYPE.get(node, "unknown_evidence"),
                    refer_to=decision.refer_to,
                    expected_keywords=self._default_expected_keywords(node, target),
                    expected_fields=self._default_expected_fields(node),
                    reason=decision.reason,
                )
            )

        return tasks

    def _default_expected_keywords(self, node: str, target: str) -> List[str]:
        if node in {AuditNodes.RELEASE_FRONT, AuditNodes.RELEASE_BACK}:
            return [target, "发版", "账号"]
        if node in {AuditNodes.APP_ADMIN, AuditNodes.APP_SCREENSHOT}:
            return [target, "应用", "账号"]
        if node == AuditNodes.APP_SERVER:
            return [target, "服务器", "账号"]
        if node in {AuditNodes.DB_SERVER, AuditNodes.DB_ACCOUNT}:
            return [target, "数据库", "账号"]
        if node == AuditNodes.DEV_LIST:
            return [target, "开发人员", "清单"]
        if node == AuditNodes.DEV_NO_PROD_ABOVE_READ:
            return [target, "开发人员", "生产", "权限"]
        if node == AuditNodes.SOD:
            return [target, "SOD", "管理员"]
        return [target]

    def _default_expected_fields(self, node: str) -> List[str]:
        if node in {AuditNodes.DB_ACCOUNT, AuditNodes.DB_SERVER, AuditNodes.APP_SERVER, AuditNodes.APP_ADMIN}:
            return ["账号名", "账号使用人", "用途", "是否在职", "校验结果"]
        if node in {AuditNodes.RELEASE_FRONT, AuditNodes.RELEASE_BACK}:
            return ["账号名", "账号使用人", "角色", "权限", "校验结果"]
        if node == AuditNodes.DEV_LIST:
            return ["开发人员姓名", "部门", "状态"]
        if node == AuditNodes.DEV_NO_PROD_ABOVE_READ:
            return ["开发人员姓名", "生产环境权限", "是否只读以上", "审批结果"]
        if node == AuditNodes.SOD:
            return ["前台管理员", "后台管理员", "是否同人", "处理结果"]
        return []


@dataclass
class OCRResult:
    sheet_name: str
    text: str
    source: str = ""
    confidence: Optional[float] = None
    warnings: List[str] = field(default_factory=list)


class OCRProviderBase:
    """OCR 接口基类。后续可接 RapidOCR / PaddleOCR / 企业内部 OCR。"""

    def extract_texts_from_workbook(self, workbook_path: str, sheet_names: Optional[List[str]] = None) -> List[OCRResult]:
        raise NotImplementedError

    def extract_text_from_image(self, image_path: str) -> OCRResult:
        raise NotImplementedError


class EvidenceCheckerBase:
    """证据检查接口基类。"""

    def check(self, task: EvidenceTask, ocr_results: List[OCRResult]) -> Dict[str, Any]:
        raise NotImplementedError


class ApprovalFlowEvidenceChecker(EvidenceCheckerBase):
    """审批流截图检查器：先做关键文本判断，审批层级规则后续再补。"""

    TIME_PATTERN = r"\d{4}-\d{1,2}-\d{1,2}\s+\d{1,2}:\d{2}:\d{2}"

    def check(self, task: EvidenceTask, ocr_results: List[OCRResult]) -> Dict[str, Any]:
        combined_text = "\n".join([r.text for r in ocr_results])
        times = re.findall(self.TIME_PATTERN, combined_text)
        return {
            "system": task.system,
            "node": task.node,
            "expected_label": task.expected_label,
            "contains_submit": "提交审阅" in combined_text,
            "contains_pass": "审阅通过" in combined_text or "通过" in combined_text,
            "contains_complete": "完成" in combined_text,
            "times": times,
            "result": "待结合审批层级规则判断",
            "raw_text": combined_text,
        }
