from __future__ import annotations

import re
from typing import Any, List, Optional, Sequence

import pandas as pd

class TextTools:
    """文本和单元格解析工具。"""

    BLANK_VALUES = {"", "nan", "none", "null", "na"}

    @staticmethod
    def is_blank(value: Any) -> bool:
        if value is None:
            return True
        try:
            if pd.isna(value):
                return True
        except TypeError:
            pass
        text = str(value).strip()
        return text.lower() in TextTools.BLANK_VALUES

    @staticmethod
    def normalize_text(value: Any) -> str:
        if TextTools.is_blank(value):
            return ""
        text = str(value).replace("\u3000", " ").strip()
        text = re.sub(r"[ \t]+", " ", text)
        return text

    @staticmethod
    def normalize_key(value: Any) -> str:
        """用于匹配 key 的标准化：去空格、统一英文大小写。"""
        text = TextTools.normalize_text(value)
        text = text.replace(" ", "")
        return text.upper()

    @staticmethod
    def split_multiline_cell(value: Any) -> List[str]:
        """
        将自动换行单元格拆成多个条目。

        例如：
        应用层服务器容器化\n无应用层用户
        -> [应用层服务器容器化, 无应用层用户]
        """
        if TextTools.is_blank(value):
            return []

        text = str(value).replace("\r\n", "\n").replace("\r", "\n")
        parts = re.split(r"\n+", text)
        cleaned = [TextTools.normalize_text(p) for p in parts]
        return [p for p in cleaned if p]

    @staticmethod
    def safe_get(row: pd.Series, col_name: str, col_index: Optional[int] = None) -> Any:
        """先按列名取值，列名不存在时按列序号取值。"""
        if col_name in row.index:
            value = row.loc[col_name]
            if isinstance(value, pd.Series):
                return value.iloc[0]
            return value

        if col_index is not None and 0 <= col_index < len(row):
            return row.iloc[col_index]

        return None

    @staticmethod
    def contains_any_regex(
        items: Sequence[str],
        positive_patterns: Sequence[str],
        negative_patterns: Optional[Sequence[str]] = None,
        flags: int = re.IGNORECASE,
    ) -> bool:
        """
        使用正则表达式判断多个文本条目中是否命中规则。

        先看 negative_patterns，避免“无应用层用户”误判为“无应用层”。
        """
        negative_patterns = negative_patterns or []

        for item in items:
            for pattern in negative_patterns:
                if re.search(pattern, item, flags=flags):
                    return False

        for item in items:
            for pattern in positive_patterns:
                if re.search(pattern, item, flags=flags):
                    return True

        return False

    @staticmethod
    def parse_full_paas(value: Any, warnings: List[str]) -> Optional[bool]:
        """
        是否完全使用 DBA PaaS 服务。

        返回：
        - True：Y，完全使用 PaaS
        - False：N 或 Partially，不能按“完全 PaaS”豁免
        - None：N/A / 空值，表示该字段不适用或无法由该字段判断
        """
        if TextTools.is_blank(value):
            warnings.append("是否使用DBA PaaS服务为空，未自动推断。")
            return None

        text = TextTools.normalize_text(value).upper()
        if text in {"Y", "YES", "TRUE", "1", "是"}:
            return True
        if text in {"N", "NO", "FALSE", "0", "否"}:
            return False
        if "PARTIALLY" in text or "部分" in text:
            return False
        if text in {"N/A", "NA", "不适用"}:
            return None

        warnings.append(f"是否使用DBA PaaS服务出现未定义值：{value!r}。")
        return None
