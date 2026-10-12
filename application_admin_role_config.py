"""应用层管理员角色完整性配置与匹配规则。"""

from __future__ import annotations

import re
from functools import lru_cache
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd

from sid_utils import normalize_sid


CONFIG_FILE = Path(__file__).resolve().parent / "configs" / "应用层管理员角色配置表.xlsx"

COLUMN_ALIASES = {
    "SID": "sid",
    "系统ID": "sid",
    "角色关键词": "role_keyword",
    "管理员角色关键词": "role_keyword",
    "管理员角色": "role_keyword",
    "匹配方式": "match_mode",
    "是否启用": "enabled",
    "备注": "note",
}
DISABLED_VALUES = {"FALSE", "F", "N", "NO", "否", "0", "禁用"}
MATCH_MODE_ALIASES = {
    "精确": "exact",
    "EXACT": "exact",
    "模糊": "fuzzy",
    "FUZZY": "fuzzy",
}


def _text(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def normalize_role(value: object) -> str:
    """角色比较用的温和标准化：保留中英文与数字，忽略分隔符和空白。"""
    return re.sub(r"[\s_\-/\\（）()【】\[\].:：]+", "", _text(value)).upper()


@lru_cache(maxsize=1)
def load_application_admin_role_rules() -> Dict[str, List[Dict[str, str]]]:
    """按 SID 读取应用层管理员角色关键词；绝不按系统名称回退匹配。"""
    if not CONFIG_FILE.exists():
        return {}

    df = pd.read_excel(CONFIG_FILE, dtype=str, keep_default_na=False).rename(columns=COLUMN_ALIASES)
    required = {"sid", "role_keyword"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"{CONFIG_FILE} 缺少字段：{sorted(missing)}")

    if "enabled" in df.columns:
        enabled = df["enabled"].astype(str).str.strip().str.upper()
        df = df[~enabled.isin(DISABLED_VALUES)]
    if "match_mode" not in df.columns:
        # 精简配置表只需维护 SID 和管理员角色关键词；默认启用保守的模糊匹配，
        # 用于覆盖“CSSC”与“用户管理-CSSC”之类的前后缀差异。
        df["match_mode"] = "模糊"
    if "note" not in df.columns:
        df["note"] = ""

    rules: Dict[str, List[Dict[str, str]]] = {}
    seen: set[Tuple[str, str]] = set()
    for _, row in df.iterrows():
        sid = _text(row["sid"])
        role_keyword = _text(row["role_keyword"])
        if not sid or not role_keyword:
            continue
        sid_key = normalize_sid(sid)
        role_key = normalize_role(role_keyword)
        duplicate_key = (sid_key, role_key)
        if duplicate_key in seen:
            raise ValueError(f"{CONFIG_FILE} 存在重复 SID/角色关键词：{sid} / {role_keyword}")
        seen.add(duplicate_key)

        match_mode_raw = _text(row.get("match_mode", "精确")).upper()
        match_mode = MATCH_MODE_ALIASES.get(match_mode_raw)
        if not match_mode:
            raise ValueError(
                f"{CONFIG_FILE} 的匹配方式仅支持 精确/模糊：SID={sid}，角色关键词={role_keyword}"
            )
        rules.setdefault(sid_key, []).append({
            "sid": sid,
            "role_keyword": role_keyword,
            "match_mode": match_mode,
            "note": _text(row.get("note", "")),
        })
    return rules


def match_role_keyword(expected: str, actual: str, match_mode: str) -> Tuple[bool, str]:
    """比较配置角色和 OCR 角色列的一个单元格，返回是否匹配及匹配方式。"""
    expected_norm = normalize_role(expected)
    actual_norm = normalize_role(actual)
    if not expected_norm or not actual_norm:
        return False, ""
    if expected_norm == actual_norm:
        return True, "精确匹配"
    if match_mode != "fuzzy":
        return False, ""

    # 覆盖“CSSC”与“用户管理-CSSC”等前后缀差异。
    if expected_norm in actual_norm or actual_norm in expected_norm:
        return True, "模糊包含匹配"

    # 对 OCR 的一两个字符误读保留容错，但短关键词避免过度匹配。
    if min(len(expected_norm), len(actual_norm)) < 4:
        return False, ""
    from difflib import SequenceMatcher
    ratio = SequenceMatcher(None, expected_norm, actual_norm).ratio()
    if ratio >= 0.80:
        return True, f"模糊相似匹配({ratio:.0%})"
    return False, ""
