from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

from audit_nodes import AuditNodes
from decisions import Decision
from field_translator import FieldFacts
from rules_release_tool import ReleaseToolRuleEngine
from text_tools import TextTools

@dataclass
class ScopeResult:
    facts: FieldFacts
    decisions: Dict[str, Decision]
    warnings: List[str]


class AccountAuditScopeEngine:
    """账号审阅范围判断引擎。"""

    # 来自 WP/Breakdown 中 Utility/Other 的特殊数据库口径。
    SYSTEM_DB_OVERRIDES: Dict[str, Tuple[Decision, Decision]] = {
        "DBAELK": (Decision.yes("WP 中 DBA ELK 数据库服务器为 是。"), Decision.yes("WP 中 DBA ELK 数据库账号为 是。")),
        "DBASOX服务器": (Decision.yes("WP 中 DBA SOX服务器数据库服务器为 是。"), Decision.na("DBA SOX服务器自身数据库账号为 N/A。")),
        "INFRAJENKINS": (Decision.na("Remark 为无数据库&数据库服务器。"), Decision.na("Remark 为无数据库&数据库服务器。")),
        "PAM": (Decision.na("WP 中 PAM 数据库服务器为 N/A。"), Decision.na("WP 中 PAM 数据库账号为 N/A。")),
        "PMS&VMS": (Decision.yes("WP 中 PMS&VMS 数据库服务器为 是。"), Decision.yes("WP 中 PMS&VMS 数据库账号为 是。")),
        "码云": (Decision.yes("WP 中码云数据库服务器为 是。"), Decision.yes("WP 中码云数据库账号为 是。")),
    }

    def __init__(self, release_rule_engine: Optional[ReleaseToolRuleEngine] = None):
        self.release_rule_engine = release_rule_engine or ReleaseToolRuleEngine()

    def determine(self, facts: FieldFacts) -> ScopeResult:
        warnings = list(facts.warnings)

        decisions: Dict[str, Decision] = {
            AuditNodes.APP_ADMIN: Decision.yes("默认需要审阅应用层管理员账号。"),
            AuditNodes.APP_SCREENSHOT: Decision.yes("默认需要上传应用层账号截图。"),
            AuditNodes.APP_SERVER: Decision.yes("默认需要审阅应用层服务器账号。"),
            AuditNodes.DB_SERVER: Decision.yes("默认需要审阅数据库服务器账号。"),
            AuditNodes.DB_ACCOUNT: Decision.yes("默认需要审阅数据库账号。"),
            AuditNodes.DEV_LIST: Decision.na("默认不需要开发人员清单，SOX/合规支持平台除外。"),
            AuditNodes.DEV_NO_PROD_ABOVE_READ: Decision.na("默认不检查开发人员生产环境权限，SOX/合规支持平台除外。"),
            AuditNodes.SOD: Decision.yes("默认检查前后台管理员 SOD。"),
        }

        # 发版工具前后台。
        release_scope = self.release_rule_engine.determine(facts)
        decisions[AuditNodes.RELEASE_FRONT] = release_scope.front
        decisions[AuditNodes.RELEASE_BACK] = release_scope.back
        warnings.extend(release_scope.warnings)

        # 无应用层。
        if facts.no_app_layer:
            decisions[AuditNodes.APP_ADMIN] = Decision.na("Remark 包含无应用层。")
            decisions[AuditNodes.APP_SCREENSHOT] = Decision.na("Remark 包含无应用层。")
            decisions[AuditNodes.APP_SERVER] = Decision.na("Remark 包含无应用层。")
            decisions[AuditNodes.SOD] = Decision.na("无应用层，不适用前后台管理员 SOD。")

        # 无应用层用户。
        if facts.no_app_layer_user:
            decisions[AuditNodes.APP_ADMIN] = Decision.na("Remark 包含应用层无用户/无应用层用户。")
            decisions[AuditNodes.APP_SCREENSHOT] = Decision.na("Remark 包含应用层无用户/无应用层用户。")
            decisions[AuditNodes.SOD] = Decision.na("无应用层用户，不适用前后台管理员 SOD。")

        # 容器化 / 实际为负载均衡服务器。
        if facts.is_app_server_containerized:
            decisions[AuditNodes.APP_SERVER] = Decision.na_with_reason("容器化无需审阅", "Remark 包含应用层/应用服务器容器化。")
        elif facts.non_app_server_or_lb:
            decisions[AuditNodes.APP_SERVER] = Decision.na("Remark 显示为负载均衡服务器或不涉及代码/数据变更。")

        # 数据库服务器 / 数据库账号。
        self._apply_database_rules(facts, decisions, warnings)

        # 开发人员清单 / 生产权限。
        if facts.is_sox_system or facts.is_compliance_support_platform:
            decisions[AuditNodes.DEV_LIST] = Decision.yes("SOX系统或合规支持平台需上传开发人员清单。")
            decisions[AuditNodes.DEV_NO_PROD_ABOVE_READ] = Decision.yes("SOX系统或合规支持平台需检查开发人员生产环境无只读以上权限。")
        elif facts.is_utility_tool or facts.is_other:
            decisions[AuditNodes.DEV_LIST] = Decision.na("Utility Tool（除合规支持平台）和 Other 不需要开发人员清单。")
            decisions[AuditNodes.DEV_NO_PROD_ABOVE_READ] = Decision.na("Utility Tool（除合规支持平台）和 Other 不检查开发人员生产权限。")

        # Utility/Other 的 SOD：先按 WP 中可见规律处理，不泛化到所有工具。
        self._apply_sod_rules(facts, decisions)

        # 无后台为强排除规则。
        if facts.no_backend:
            decisions[AuditNodes.RELEASE_FRONT] = Decision.na("Remark 包含无后台。")
            decisions[AuditNodes.RELEASE_BACK] = Decision.na("Remark 包含无后台。")
            decisions[AuditNodes.APP_SERVER] = Decision.na("Remark 包含无后台。")
            decisions[AuditNodes.DB_SERVER] = Decision.na("Remark 包含无后台。")
            decisions[AuditNodes.DB_ACCOUNT] = Decision.na("Remark 包含无后台。")
            decisions[AuditNodes.SOD] = Decision.na("Remark 包含无后台。")

        return ScopeResult(facts=facts, decisions=decisions, warnings=warnings)

    def _apply_database_rules(self, facts: FieldFacts, decisions: Dict[str, Decision], warnings: List[str]) -> None:
        system_key = TextTools.normalize_key(facts.system)

        if facts.no_database_or_db_server:
            decisions[AuditNodes.DB_SERVER] = Decision.na("Remark 包含无数据库/无数据库服务器。")
            decisions[AuditNodes.DB_ACCOUNT] = Decision.na("Remark 包含无数据库/无数据库服务器。")
            return

        if system_key in self.SYSTEM_DB_OVERRIDES:
            db_server, db_account = self.SYSTEM_DB_OVERRIDES[system_key]
            decisions[AuditNodes.DB_SERVER] = db_server
            decisions[AuditNodes.DB_ACCOUNT] = db_account
            return

        if facts.uses_full_dba_paas is True:
            decisions[AuditNodes.DB_SERVER] = Decision.refer("DBA SOX服务器", "完全使用 DBA PaaS 服务，数据库服务器账号 Refer to DBA SOX服务器。")
            decisions[AuditNodes.DB_ACCOUNT] = Decision.yes("PaaS 仅影响数据库服务器账号，数据库账号仍需审阅。")
        elif facts.uses_full_dba_paas is False:
            decisions[AuditNodes.DB_SERVER] = Decision.yes("未完全使用 DBA PaaS 服务，数据库服务器账号需审阅。")
            decisions[AuditNodes.DB_ACCOUNT] = Decision.yes("数据库账号需审阅。")
        else:
            if facts.is_utility_tool or facts.is_other:
                decisions[AuditNodes.DB_SERVER] = Decision.unknown("PaaS=N/A 且非系统级数据库覆盖项，是否审阅数据库服务器需确认。")
                decisions[AuditNodes.DB_ACCOUNT] = Decision.unknown("PaaS=N/A 且非系统级数据库覆盖项，是否审阅数据库账号需确认。")
                warnings.append(f"{facts.system!r} 的数据库/数据库服务器口径需要确认。")

    def _apply_sod_rules(self, facts: FieldFacts, decisions: Dict[str, Decision]) -> None:
        system_key = TextTools.normalize_key(facts.system)

        if facts.no_app_layer or facts.no_app_layer_user or facts.no_backend:
            decisions[AuditNodes.SOD] = Decision.na("无应用层/无应用层用户/无后台，不适用前后台管理员 SOD。")
            return

        if facts.is_sox_system or facts.is_compliance_support_platform:
            decisions[AuditNodes.SOD] = Decision.yes("SOX系统或合规支持平台默认检查前后台管理员 SOD。")
            return

        if system_key in {"INFRAJENKINS", "PAM", "PMS&VMS", "码云"}:
            decisions[AuditNodes.SOD] = Decision.yes("WP 中该 Utility Tool 需要检查 SOD。")
        elif facts.is_utility_tool or facts.is_other:
            decisions[AuditNodes.SOD] = Decision.na("Utility/Other 未命中需检查 SOD 的系统覆盖项。")
