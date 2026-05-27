from __future__ import annotations

from typing import Any, Dict, List, Optional

import pandas as pd

from audit_nodes import AuditNodes
from config import AuditConfig
from decisions import Decision
from field_translator import FieldTranslator
from rules_scope import AccountAuditScopeEngine

class AccountAuditPipeline:
    """一键执行入口。"""

    def __init__(self, config: Optional[AuditConfig] = None):
        self.config = config or AuditConfig()
        self.translator = FieldTranslator(self.config)
        self.scope_engine = AccountAuditScopeEngine()

    def run(self, df: pd.DataFrame) -> pd.DataFrame:
        rows: List[Dict[str, Any]] = []

        for i, (_, row) in enumerate(df.iterrows(), start=2):
            facts = self.translator.translate_row(row, row_number=i)
            result = self.scope_engine.determine(facts)

            output: Dict[str, Any] = {
                "Excel行号": facts.row_number,
                "Category": facts.category,
                "System": facts.system,
                "服务器类型": facts.server_type,
                "PMS流程_拆分后": " | ".join(facts.pms_flow_items),
                "发版工具_拆分后": " | ".join(facts.release_tool_items),
                "Remark_拆分后": " | ".join(facts.remark_items),
                "uses_full_dba_paas": facts.uses_full_dba_paas,
                "is_app_server_containerized": facts.is_app_server_containerized,
                "no_app_layer": facts.no_app_layer,
                "no_app_layer_user": facts.no_app_layer_user,
                "no_backend": facts.no_backend,
                "no_database_or_db_server": facts.no_database_or_db_server,
                "non_app_server_or_lb": facts.non_app_server_or_lb,
                "is_sox_system": facts.is_sox_system,
                "is_compliance_support_platform": facts.is_compliance_support_platform,
                "is_utility_tool": facts.is_utility_tool,
                "is_other": facts.is_other,
                "is_pms_vms": facts.is_pms_vms,
            }

            for node in AuditNodes.ALL:
                decision = result.decisions.get(node, Decision.unknown("节点无判断结果。"))
                output[node] = decision.label
                output[f"{node}_是否需凭证"] = "是" if decision.requires_evidence else "否"
                output[f"{node}_ReferTo"] = decision.refer_to or ""
                output[f"{node}_原因"] = decision.reason

            output["Warnings"] = " | ".join(result.warnings)
            rows.append(output)

        return pd.DataFrame(rows)

    @staticmethod
    def read_file(file_path: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
        lower = file_path.lower()
        if lower.endswith((".xlsx", ".xls")):
            return pd.read_excel(file_path, sheet_name=sheet_name or 0)
        if lower.endswith(".csv"):
            return pd.read_csv(file_path)
        raise ValueError("当前只支持 .xlsx / .xls / .csv 文件。")
