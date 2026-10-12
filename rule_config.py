from __future__ import annotations

from functools import lru_cache
from pathlib import Path
from typing import Dict

import pandas as pd

from decisions import Decision
from sid_utils import normalize_sid


CONFIG_DIR = Path(__file__).resolve().parent / "configs"
SYSTEM_RULES_FILE = CONFIG_DIR / "系统规则特殊覆盖表.xlsx"
LEGACY_SYSTEM_RULES_FILE = CONFIG_DIR / "system_rule_overrides.csv"

SYSTEM_RULE_COLUMN_ALIASES = {
    "SID": "sid",
    "系统ID": "sid",
    "是否合规支持平台": "is_compliance_support_platform",
    "是否PMS_VMS": "is_pms_vms",
    "数据库服务器结论": "db_server_label",
    "数据库服务器原因": "db_server_reason",
    "数据库账号结论": "db_account_label",
    "数据库账号原因": "db_account_reason",
    "是否检查SOD": "sod_required",
    "SOD原因": "sod_reason",
    "发版工具前台结论": "release_front_label",
    "发版工具前台原因": "release_front_reason",
    "发版工具后台结论": "release_back_label",
    "发版工具后台原因": "release_back_reason",
    "警告提示": "warning",
}


def _text(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


@lru_cache(maxsize=1)
def load_system_rule_overrides() -> Dict[str, Dict[str, str]]:
    """读取系统级业务特例；配置缺失或字段错误时立即报错，避免静默套用错误规则。"""
    if SYSTEM_RULES_FILE.exists():
        config_file = SYSTEM_RULES_FILE
        df = pd.read_excel(config_file, dtype=str, keep_default_na=False)
    elif LEGACY_SYSTEM_RULES_FILE.exists():
        config_file = LEGACY_SYSTEM_RULES_FILE
        df = pd.read_csv(config_file, dtype=str, keep_default_na=False)
    else:
        raise FileNotFoundError(f"系统规则配置不存在：{SYSTEM_RULES_FILE}")

    df = df.rename(columns=SYSTEM_RULE_COLUMN_ALIASES)
    required = {
        "sid",
        "is_compliance_support_platform",
        "is_pms_vms",
        "db_server_label",
        "db_server_reason",
        "db_account_label",
        "db_account_reason",
        "sod_required",
        "sod_reason",
        "release_front_label",
        "release_front_reason",
        "release_back_label",
        "release_back_reason",
        "warning",
    }
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"{config_file} 缺少字段：{sorted(missing)}")

    if "enabled" in df.columns:
        disabled = {"FALSE", "F", "N", "NO", "否", "0", "禁用"}
        enabled_text = df["enabled"].astype(str).str.strip().str.upper()
        df = df[~enabled_text.isin(disabled)]
    if "是否启用" in df.columns:
        disabled = {"FALSE", "F", "N", "NO", "否", "0", "禁用"}
        enabled_text = df["是否启用"].astype(str).str.strip().str.upper()
        df = df[~enabled_text.isin(disabled)]

    rules: Dict[str, Dict[str, str]] = {}
    for _, row in df.iterrows():
        key = normalize_sid(row["sid"])
        if key:
            if key in rules:
                raise ValueError(f"{config_file} 存在重复 SID：{row['sid']}")
            rules[key] = {column: _text(row[column]) for column in required}
    return rules


def decision_from_label(label: str, reason: str = "") -> Decision:
    """把配置表中的业务标签恢复为 Decision。"""
    label = _text(label)
    if label == "是":
        return Decision.yes(reason)
    if label == "N/A":
        return Decision.na(reason)
    if label.startswith("Refer to "):
        return Decision.refer(label.removeprefix("Refer to ").strip(), reason)
    if label.startswith("N/A-Refer to "):
        return Decision.na_refer(label.removeprefix("N/A-Refer to ").strip(), reason)
    if label.startswith("N/A-"):
        return Decision.custom(label, requires_evidence=False, reason=reason)
    if not label:
        return Decision.unknown(reason or "配置表中的结论为空。")
    return Decision.custom(label, requires_evidence=True, reason=reason)
