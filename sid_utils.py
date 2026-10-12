from __future__ import annotations

import re
from typing import Any

import pandas as pd


def normalize_sid(value: Any) -> str:
    """把 Excel/CSV 中的 SID 统一为可精确比较的文本键。"""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass

    text = str(value).strip()
    if not text:
        return ""

    # SID 列含空值时，pandas 可能把纯数字 SID 读成 1263.0。
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".", 1)[0]
    return text.replace(" ", "").upper()

