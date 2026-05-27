from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Optional

import pandas as pd

from config import AuditConfig
from text_tools import TextTools

@dataclass
class FieldFacts:
    row_number: Optional[int]

    category: str
    system: str
    server_type: str
    pms_flow_raw: str
    release_tool_raw: str
    remark_raw: str

    pms_flow_items: List[str]
    release_tool_items: List[str]
    remark_items: List[str]

    uses_full_dba_paas: Optional[bool]
    is_app_server_containerized: bool
    no_app_layer: bool
    no_app_layer_user: bool
    no_backend: bool
    no_database_or_db_server: bool
    non_app_server_or_lb: bool

    is_sox_system: bool
    is_compliance_support_platform: bool
    is_utility_tool: bool
    is_other: bool
    is_pms_vms: bool

    warnings: List[str] = field(default_factory=list)


class FieldTranslator:
    """把 DataFrame 一行翻译成 FieldFacts。"""

    def __init__(self, config: Optional[AuditConfig] = None):
        self.config = config or AuditConfig()

    def translate_row(self, row: pd.Series, row_number: Optional[int] = None) -> FieldFacts:
        cfg = self.config
        warnings: List[str] = []

        category_raw = TextTools.safe_get(row, cfg.category_col, cfg.category_col_index)
        system_raw = TextTools.safe_get(row, cfg.system_col, cfg.system_col_index)
        paas_raw = TextTools.safe_get(row, cfg.paas_col, cfg.paas_col_index)
        server_type_raw = TextTools.safe_get(row, cfg.server_type_col, cfg.server_type_col_index)
        pms_flow_raw = TextTools.safe_get(row, cfg.pms_flow_col, cfg.pms_flow_col_index)
        release_tool_raw = TextTools.safe_get(row, cfg.release_tool_col, cfg.release_tool_col_index)
        remark_raw = TextTools.safe_get(row, cfg.remark_col, cfg.remark_col_index)

        category = TextTools.normalize_text(category_raw)
        system = TextTools.normalize_text(system_raw)
        server_type = TextTools.normalize_text(server_type_raw)
        pms_flow_text = TextTools.normalize_text(pms_flow_raw)
        release_tool_text = TextTools.normalize_text(release_tool_raw)
        remark_text = TextTools.normalize_text(remark_raw)

        pms_flow_items = TextTools.split_multiline_cell(pms_flow_raw)
        release_tool_items = TextTools.split_multiline_cell(release_tool_raw)
        remark_items = TextTools.split_multiline_cell(remark_raw)

        uses_full_dba_paas = TextTools.parse_full_paas(paas_raw, warnings)

        is_app_server_containerized = TextTools.contains_any_regex(
            remark_items,
            positive_patterns=[r"应用\s*(?:层)?\s*服务器.*容器化"],
            negative_patterns=[
                r"未\s*容器化",
                r"非\s*容器化",
                r"不\s*涉及\s*容器化",
                r"没有\s*容器化",
            ],
        )

        no_app_layer = TextTools.contains_any_regex(
            remark_items,
            positive_patterns=[r"无\s*应用层"],
            negative_patterns=[r"无\s*应用层\s*用户", r"应用层\s*无\s*用户"],
        )

        no_app_layer_user = TextTools.contains_any_regex(
            remark_items,
            positive_patterns=[r"应用层\s*无\s*用户", r"无\s*应用层\s*用户"],
        )

        no_backend = TextTools.contains_any_regex(
            remark_items,
            positive_patterns=[r"无\s*后台"],
        )

        no_database_or_db_server = TextTools.contains_any_regex(
            remark_items,
            positive_patterns=[r"无\s*数据库\s*&?\s*数据库服务器", r"无\s*数据库服务器", r"无\s*数据库"],
        )

        non_app_server_or_lb = TextTools.contains_any_regex(
            remark_items,
            positive_patterns=[r"负载均衡服务器.*不涉及代码/数据变更", r"不涉及代码/数据变更"],
        )

        category_key = TextTools.normalize_key(category)
        system_key = TextTools.normalize_key(system)

        is_sox_system = category_key in {"SOXSYSTEM", "SOX系统"}
        is_compliance_support_platform = system_key == "合规支持系统"
        is_utility_tool = category_key == "UTILITYTOOL"
        is_other = category_key in {"OTHER", "其他"}
        is_pms_vms = system_key in {"PMS&VMS", "PMS/VMS"}

        return FieldFacts(
            row_number=row_number,
            category=category,
            system=system,
            server_type=server_type,
            pms_flow_raw=pms_flow_text,
            release_tool_raw=release_tool_text,
            remark_raw=remark_text,
            pms_flow_items=pms_flow_items,
            release_tool_items=release_tool_items,
            remark_items=remark_items,
            uses_full_dba_paas=uses_full_dba_paas,
            is_app_server_containerized=is_app_server_containerized,
            no_app_layer=no_app_layer,
            no_app_layer_user=no_app_layer_user,
            no_backend=no_backend,
            no_database_or_db_server=no_database_or_db_server,
            non_app_server_or_lb=non_app_server_or_lb,
            is_sox_system=is_sox_system,
            is_compliance_support_platform=is_compliance_support_platform,
            is_utility_tool=is_utility_tool,
            is_other=is_other,
            is_pms_vms=is_pms_vms,
            warnings=warnings,
        )
