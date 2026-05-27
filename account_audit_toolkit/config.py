from __future__ import annotations

from dataclasses import dataclass

@dataclass
class AuditConfig:
    """
    字段配置。

    默认字段来自 WP：
    A列 Category
    B列 System
    D列 是否使用DBA PaaS服务
    F列 PMS流程
    G列 发版工具
    H列 Remark
    """

    category_col: str = "Category"
    system_col: str = "System"
    paas_col: str = "是否使用DBA PaaS服务"
    server_type_col: str = "服务器类型"
    pms_flow_col: str = "PMS流程"
    release_tool_col: str = "发版工具"
    remark_col: str = "Remark"

    category_col_index: int = 0
    system_col_index: int = 1
    paas_col_index: int = 3
    server_type_col_index: int = 4
    pms_flow_col_index: int = 5
    release_tool_col_index: int = 6
    remark_col_index: int = 7
