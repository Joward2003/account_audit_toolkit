from __future__ import annotations

"""
execution_plan_builder.py

用途：
在已经生成 todo_list_with_folder.xlsx 之后，继续生成“执行计划”。

它不做 OCR，不直接下审计结论，而是先把后续检查拆成 3 个 sheet：

1. expected_accounts
   从每个系统文件夹里的“账号审阅导出”Excel 中，根据 node 对应的“审阅对象”
   筛出理论上应该检查的账号。默认取：
   - 审阅对象
   - 账号名
   - 账号使用人
   - 用途

2. evidence_plan
   判断每条任务后续需要什么证据：
   - 是否需要手工 supporting / 截图
   - 证据格式是什么
   - 后续应该用 OCR、Excel 读取、导出可用性检查，还是派生集合判断

3. compare_result
   先生成比对结果占位。
   对“只需要账号审阅导出即可判断”的节点，先给出 PASS_EXPORT_AVAILABLE / REVIEW。
   对需要 OCR 的节点，先标记 PENDING_OCR，不在这里强行判断。

当前设计原则：
- Scope Engine 只回答“哪些 node 理论上要查”。
- 本模块回答“进入检查范围的 node，具体怎么执行”。
- 账号审阅导出是 expected set 来源。
- supporting.xlsx / 截图 / Excel 是 actual set 来源，留给后续 OCR / Evidence Checker。
- SYSTEM / 系统账号 / 服务账号默认不进入 expected set，但会留痕。
"""

import argparse
import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import pandas as pd


# ============================================================
# 1. 固定规则：Node -> 账号审阅导出的审阅对象
# ============================================================

# 这里刻意写进代码，不放 Excel。
# 因为这些是底层业务映射，变动频率低，且需要稳定。
NODE_TO_REVIEW_OBJECT: Dict[str, str] = {
    "需审阅发版工具发版账号": "发布工具",
    "需审阅发版工具后台账号": "发布工具服务器",
    "需审阅应用层管理员账号": "应用层",
    "需审阅应用层服务器账号": "应用服务器",
    "需审阅数据库服务器账号": "数据库服务器",
    "需审阅数据库账号": "数据库",
}

# 某些 node 本质上不是“再做一次账号比对”，而是证明另一项审阅有 supporting。
# 例如：需上传应用层账号截图 = 支撑“需审阅应用层管理员账号”的证据完整性。
NODE_LINKED_TO: Dict[str, str] = {
    "需上传应用层账号截图": "需审阅应用层管理员账号",
}

# Evidence-only node 不自己生成 expected account set。
# 但它仍然需要借用关联审阅对象的“手工上传标识”判断材料是否已上传。
EVIDENCE_ONLY_NODE_REVIEW_OBJECT: Dict[str, str] = {
    "需上传应用层账号截图": "应用层",
}

# 默认 supporting.xlsx sheet 名。后续如果你已经有 supporting_sheet_mapping.xlsx，
# 可以再改成从配置表读取。
DEFAULT_SUPPORTING_SHEET: Dict[str, str] = {
    "需审阅发版工具发版账号": "发版工具发版账号",
    "需审阅发版工具后台账号": "发版工具后台账号",
    "需审阅应用层管理员账号": "应用层管理员账号",
    "需上传应用层账号截图": "应用层管理员账号",
    "需审阅应用层服务器账号": "应用层服务器账号",
    "需审阅数据库服务器账号": "数据库服务器账号",
    "需审阅数据库账号": "数据库账号",
    "需上传开发人员清单": "开发人员清单",
    "开发人员在生产环境无只读以上权限": "开发人员权限检查",
    "无前后台管理员SOD问题": "SOD检查",
}

# 默认读取账号审阅导出中的这几列。
EXPORT_REVIEW_OBJECT_CANDIDATES = ["审阅对象", "审阅对象名称", "ReviewObject", "review_object"]
EXPORT_ACCOUNT_CANDIDATES = ["账号名", "账号", "用户账号", "User/group", "User", "Account", "account"]
EXPORT_OWNER_CANDIDATES = ["账号使用人", "使用人", "用户", "用户名", "UserName", "Owner", "owner"]
EXPORT_USAGE_CANDIDATES = ["用途", "身份", "权限", "角色", "Role", "Usage", "usage"]
EXPORT_SERVER_ADDRESS_CANDIDATES = ["服务器地址", "服务器", "Server", "server", "主机名", "Hostname"]
EXPORT_MANUAL_FLAG_CANDIDATES = ["手工上传标识", "是否手工上传", "ManualUploadFlag"]
EXPORT_CHECK_RESULT_CANDIDATES = ["校验结果", "检查结果", "CheckResult"]


# ============================================================
# 2. 通用工具函数
# ============================================================

def is_blank(value: Any) -> bool:
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except TypeError:
        pass
    return str(value).strip() == ""


def clean_text(value: Any) -> str:
    if is_blank(value):
        return ""
    text = str(value).replace("\u3000", " ").strip()
    text = re.sub(r"[ \t]+", " ", text)
    return text


def normalize_key(value: Any) -> str:
    """
    用于匹配字段名、账号名、服务器类型的弱标准化。
    """
    text = clean_text(value).upper()
    text = re.sub(r"[\s_\-/\\（）()【】\[\].:：]+", "", text)
    return text


def normalize_account(value: Any) -> str:
    """
    第一版账号标准化：
    - 去空
    - 统一大写
    - 去掉空格、下划线、中划线、斜杠等常见 OCR 干扰符号

    注意：
    Jenkins 这类角色矩阵中的 user/group 可能是英文名，不一定等于账号。
    第一版只做 presence check，不做最终身份认定。
    """
    return normalize_key(value)


def split_multi_values(value: Any) -> List[str]:
    if is_blank(value):
        return []
    parts = re.split(r"[;；|,，\n\r]+", str(value))
    return [clean_text(x) for x in parts if clean_text(x)]


def resolve_column(df: pd.DataFrame, candidates: Iterable[str]) -> Optional[str]:
    """
    在 df 中查找候选列名。
    先精确匹配，再标准化匹配。
    """
    columns = list(df.columns)

    # exact
    for cand in candidates:
        if cand in columns:
            return cand

    norm_to_col = {normalize_key(c): c for c in columns}
    for cand in candidates:
        key = normalize_key(cand)
        if key in norm_to_col:
            return norm_to_col[key]

    return None


def normalize_server_type(value: Any) -> str:
    """
    标准化 WP 里的“服务器类型”。

    返回：
    - Windows
    - Linux
    - Mixed
    - Unknown
    """
    text = clean_text(value).upper()
    if not text:
        return "Unknown"

    has_windows = bool(re.search(r"WINDOWS|WIN\b", text, flags=re.IGNORECASE))
    has_linux = bool(re.search(r"LINUX|LINUX/UNIX|UNIX", text, flags=re.IGNORECASE))

    if has_windows and has_linux:
        return "Mixed"
    if has_windows:
        return "Windows"
    if has_linux:
        return "Linux"
    return "Unknown"


def truthy_text(value: Any) -> bool:
    """
    判断 Excel 里的 TRUE/Y/是 等。
    """
    if is_blank(value):
        return False
    text = clean_text(value).upper()
    return text in {"TRUE", "T", "Y", "YES", "是", "1", "启用"}


def falsey_text(value: Any) -> bool:
    if is_blank(value):
        return False
    text = clean_text(value).upper()
    return text in {"FALSE", "F", "N", "NO", "否", "0", "禁用"}


MANUAL_UPLOAD_Y_VALUES = {"Y", "YES", "TRUE", "T", "1", "是", "已上传", "上传"}
MANUAL_UPLOAD_N_VALUES = {"N", "NO", "FALSE", "F", "0", "否", "未上传", "没上传"}


def normalize_manual_upload_flag(value: Any) -> str:
    """
    标准化账号审阅导出里的“手工上传标识”。

    返回：
    - Y：已手工上传材料 / 截图
    - N：未手工上传材料 / 截图
    - BLANK：空值
    - OTHER:xxx：出现未定义值，后续人工确认
    """
    if is_blank(value):
        return "BLANK"
    text = clean_text(value).upper()
    if text in MANUAL_UPLOAD_Y_VALUES:
        return "Y"
    if text in MANUAL_UPLOAD_N_VALUES:
        return "N"
    return f"OTHER:{clean_text(value)}"


def summarize_manual_upload_flags(filtered_df: pd.DataFrame, manual_flag_col: Optional[str]) -> Dict[str, Any]:
    """
    汇总某个审阅对象下的手工上传标识。

    注意：
    这里不再靠复杂服务器规则判断是否上传截图，而是直接相信账号审阅导出里的“手工上传标识”。
    """
    if manual_flag_col is None:
        return {
            "ManualUploadColumn": "",
            "ManualUploadYCount": 0,
            "ManualUploadNCount": 0,
            "ManualUploadBlankCount": 0,
            "ManualUploadOtherValues": "",
            "ManualUploadOverall": "UNKNOWN",
            "ManualUploadStatus": "MANUAL_UPLOAD_COL_MISSING",
            "ManualUploadNote": "账号审阅导出缺少 手工上传标识 列，无法判断材料是否上传。",
        }

    flags = [normalize_manual_upload_flag(v) for v in filtered_df[manual_flag_col].tolist()]
    y_count = sum(1 for x in flags if x == "Y")
    n_count = sum(1 for x in flags if x == "N")
    blank_count = sum(1 for x in flags if x == "BLANK")
    other_values = sorted({x.replace("OTHER:", "", 1) for x in flags if x.startswith("OTHER:")})

    if y_count > 0 and n_count == 0 and not other_values:
        overall = "Y"
        status = "MANUAL_UPLOAD_Y"
        note = "手工上传标识均为 Y，后续可进入 evidence / OCR 检查。"
    elif y_count == 0 and n_count > 0 and not other_values:
        overall = "N"
        status = "MANUAL_UPLOAD_N"
        note = "手工上传标识均为 N，表示未手工上传材料。"
    elif y_count > 0 and n_count > 0:
        overall = "PARTIAL"
        status = "MANUAL_UPLOAD_MIXED_Y_N"
        note = "手工上传标识同时存在 Y 和 N，需要按行或人工确认。"
    elif y_count > 0:
        overall = "PARTIAL"
        status = "MANUAL_UPLOAD_Y_WITH_BLANK_OR_OTHER"
        note = "存在 Y，但也有空值或未定义值，建议人工复核。"
    elif n_count > 0:
        overall = "N_OR_REVIEW"
        status = "MANUAL_UPLOAD_N_WITH_BLANK_OR_OTHER"
        note = "存在 N，同时有空值或未定义值，建议人工确认。"
    elif blank_count > 0 and not other_values:
        overall = "UNKNOWN"
        status = "MANUAL_UPLOAD_BLANK"
        note = "手工上传标识为空，无法判断是否上传。"
    else:
        overall = "UNKNOWN"
        status = "MANUAL_UPLOAD_OTHER_VALUE"
        note = "手工上传标识存在未定义值，需要人工确认。"

    return {
        "ManualUploadColumn": manual_flag_col,
        "ManualUploadYCount": y_count,
        "ManualUploadNCount": n_count,
        "ManualUploadBlankCount": blank_count,
        "ManualUploadOtherValues": " | ".join(other_values),
        "ManualUploadOverall": overall,
        "ManualUploadStatus": status,
        "ManualUploadNote": note,
    }


# ============================================================
# 3. 读取账号审阅导出
# ============================================================

@dataclass
class AccountExportReadResult:
    status: str
    file_path: str = ""
    sheet_name: str = ""
    df: Optional[pd.DataFrame] = None
    message: str = ""


def find_account_export_file(folder_path: str | Path) -> Optional[Path]:
    """
    在系统文件夹里寻找“账号审阅导出”Excel。

    策略：
    1. 先按文件名包含「账号审阅导出」匹配（处理 UTF-8 编码正常的场景）
    2. 文件名匹配失败时（如 zip 编码损坏），按内容匹配：
       扫描所有 xlsx，找到包含「审阅对象」和「账号名」列的文件
    """
    folder = Path(folder_path)
    if not folder.exists():
        return None

    # 收集所有 xlsx 文件
    all_xlsx: List[Path] = []
    for p in folder.rglob("*"):
        if not p.is_file():
            continue
        if p.name.startswith("~$"):
            continue
        if p.suffix.lower() not in {".xlsx", ".xls"}:
            continue
        all_xlsx.append(p)

    if not all_xlsx:
        return None

    # ---- 第 1 轮：文件名匹配 ----
    name_matches = [p for p in all_xlsx if "账号审阅导出" in p.name]
    if name_matches:
        candidates = sorted(name_matches, key=lambda x: (len(x.parts), len(x.name), x.name))
        return candidates[0]

    # ---- 第 2 轮：内容匹配（文件名编码损坏时的兜底） ----
    for p in sorted(all_xlsx, key=lambda x: (len(x.parts), len(x.name), x.name)):
        try:
            sheets = pd.read_excel(p, sheet_name=None)
            for _sname, df in sheets.items():
                review_col = resolve_column(df, EXPORT_REVIEW_OBJECT_CANDIDATES)
                account_col = resolve_column(df, EXPORT_ACCOUNT_CANDIDATES)
                if review_col and account_col:
                    return p
        except Exception:
            continue

    return None


def read_account_export(export_file: Path, sheet_name: Optional[str] = None) -> AccountExportReadResult:
    """
    读取账号审阅导出 Excel。

    如果没有指定 sheet_name，则扫描所有 sheet，
    找到同时包含“审阅对象”和“账号名”候选列的 sheet。
    """
    try:
        if sheet_name:
            df = pd.read_excel(export_file, sheet_name=sheet_name)
            review_col = resolve_column(df, EXPORT_REVIEW_OBJECT_CANDIDATES)
            account_col = resolve_column(df, EXPORT_ACCOUNT_CANDIDATES)
            if not review_col or not account_col:
                return AccountExportReadResult(
                    status="EXPORT_REQUIRED_COLUMNS_MISSING",
                    file_path=str(export_file),
                    sheet_name=str(sheet_name),
                    df=df,
                    message="指定 sheet 缺少 审阅对象 或 账号名 候选列。",
                )
            return AccountExportReadResult(
                status="OK",
                file_path=str(export_file),
                sheet_name=str(sheet_name),
                df=df,
            )

        sheets = pd.read_excel(export_file, sheet_name=None)
        fallback_first: Optional[Tuple[str, pd.DataFrame]] = None

        for sname, df in sheets.items():
            if fallback_first is None:
                fallback_first = (sname, df)

            review_col = resolve_column(df, EXPORT_REVIEW_OBJECT_CANDIDATES)
            account_col = resolve_column(df, EXPORT_ACCOUNT_CANDIDATES)
            if review_col and account_col:
                return AccountExportReadResult(
                    status="OK",
                    file_path=str(export_file),
                    sheet_name=str(sname),
                    df=df,
                )

        if fallback_first is not None:
            sname, df = fallback_first
            return AccountExportReadResult(
                status="EXPORT_REQUIRED_COLUMNS_MISSING",
                file_path=str(export_file),
                sheet_name=str(sname),
                df=df,
                message="未找到同时包含 审阅对象 和 账号名 的 sheet。",
            )

        return AccountExportReadResult(
            status="EXPORT_EMPTY_WORKBOOK",
            file_path=str(export_file),
            message="Excel 没有可读取的 sheet。",
        )

    except Exception as exc:
        return AccountExportReadResult(
            status="EXPORT_READ_ERROR",
            file_path=str(export_file),
            message=str(exc),
        )


# ============================================================
# 4. Expected accounts 构建
# ============================================================

def is_system_account_row(row: pd.Series, account_col: Optional[str], owner_col: Optional[str], usage_col: Optional[str]) -> Tuple[bool, str]:
    """
    系统账号过滤已移除——所有账号均纳入 expected set。

    保留此函数以维持接口兼容性，始终返回 False。
    """
    return False, ""


def build_expected_accounts_for_task(
    export_result: AccountExportReadResult,
    task: pd.Series,
    review_object: str,
) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """
    针对单条任务，从账号审阅导出中筛 expected accounts。

    返回：
    - expected_rows: 一账号一行，用于 expected_accounts sheet
    - summary: 任务层面的汇总，用于 evidence_plan / compare_result
    """
    task_id = clean_text(task.get("TaskID", ""))
    system = clean_text(task.get("System", ""))
    target_system = clean_text(task.get("TargetSystem", system))
    node = clean_text(task.get("检查节点", ""))
    server_type = clean_text(task.get("服务器类型", ""))
    refer_to = clean_text(task.get("ReferTo", ""))
    scope_label = clean_text(task.get("Scope结果", ""))

    base_summary = {
        "TaskID": task_id,
        "System": system,
        "TargetSystem": target_system,
        "Node": node,
        "ReviewObject": review_object,
        "ServerType": server_type,
        "ReferTo": refer_to,
        "ScopeLabel": scope_label,
        "ExportFile": export_result.file_path,
        "ExportSheet": export_result.sheet_name,
        "RawMatchedRows": 0,
        "ExpectedCount": 0,
        "SkippedCount": 0,
        "ExpectedAccounts": "",
        "ExpectedSetStatus": "",
        "ExpectedSetNote": "",
        "ManualUploadColumn": "",
        "ManualUploadYCount": 0,
        "ManualUploadNCount": 0,
        "ManualUploadBlankCount": 0,
        "ManualUploadOtherValues": "",
        "ManualUploadOverall": "UNKNOWN",
        "ManualUploadStatus": "NOT_EVALUATED",
        "ManualUploadNote": "",
    }

    if not review_object:
        base_summary["ExpectedSetStatus"] = "NO_REVIEW_OBJECT_RULE"
        base_summary["ExpectedSetNote"] = "该 node 不通过账号审阅导出筛 expected accounts，可能是 evidence_exists 或 derived check。"
        return [], base_summary

    if export_result.status != "OK" or export_result.df is None:
        base_summary["ExpectedSetStatus"] = export_result.status
        base_summary["ExpectedSetNote"] = export_result.message or "账号审阅导出不可用。"
        return [], base_summary

    df = export_result.df.copy()

    review_col = resolve_column(df, EXPORT_REVIEW_OBJECT_CANDIDATES)
    account_col = resolve_column(df, EXPORT_ACCOUNT_CANDIDATES)
    owner_col = resolve_column(df, EXPORT_OWNER_CANDIDATES)
    usage_col = resolve_column(df, EXPORT_USAGE_CANDIDATES)
    server_addr_col = resolve_column(df, EXPORT_SERVER_ADDRESS_CANDIDATES)
    manual_flag_col = resolve_column(df, EXPORT_MANUAL_FLAG_CANDIDATES)
    check_result_col = resolve_column(df, EXPORT_CHECK_RESULT_CANDIDATES)

    if not review_col:
        base_summary["ExpectedSetStatus"] = "EXPORT_MISSING_REVIEW_OBJECT_COL"
        base_summary["ExpectedSetNote"] = "账号审阅导出缺少 审阅对象 列。"
        return [], base_summary

    if not account_col:
        base_summary["ExpectedSetStatus"] = "EXPORT_MISSING_ACCOUNT_COL"
        base_summary["ExpectedSetNote"] = "账号审阅导出缺少 账号名 列。"
        return [], base_summary

    filtered = df[df[review_col].astype(str).map(clean_text) == review_object].copy()
    base_summary["RawMatchedRows"] = len(filtered)

    if filtered.empty:
        base_summary["ExpectedSetStatus"] = "REVIEW_OBJECT_NOT_FOUND"
        base_summary["ExpectedSetNote"] = f"账号审阅导出中没有 审阅对象 = {review_object} 的记录。"
        return [], base_summary

    manual_summary = summarize_manual_upload_flags(filtered, manual_flag_col)
    base_summary.update(manual_summary)

    expected_rows: List[Dict[str, Any]] = []
    valid_accounts: List[str] = []
    skipped_count = 0

    for idx, row in filtered.iterrows():
        account_raw = clean_text(row.get(account_col, ""))
        account_std = normalize_account(account_raw)

        owner = clean_text(row.get(owner_col, "")) if owner_col else ""
        usage = clean_text(row.get(usage_col, "")) if usage_col else ""
        server_addr = clean_text(row.get(server_addr_col, "")) if server_addr_col else ""
        manual_flag = clean_text(row.get(manual_flag_col, "")) if manual_flag_col else ""
        check_result = clean_text(row.get(check_result_col, "")) if check_result_col else ""

        is_system, skip_reason = is_system_account_row(row, account_col, owner_col, usage_col)

        if not account_std:
            expected_status = "SKIP"
            skip_reason = "BLANK_ACCOUNT"
            skipped_count += 1
        else:
            expected_status = "KEEP"
            valid_accounts.append(account_std)

        expected_rows.append(
            {
                "TaskID": task_id,
                "System": system,
                "TargetSystem": target_system,
                "Node": node,
                "ReviewObject": review_object,
                "ExportFile": export_result.file_path,
                "ExportSheet": export_result.sheet_name,
                "SourceRowIndex": int(idx) + 2,
                "ReviewObjectValue": clean_text(row.get(review_col, "")),
                "AccountRaw": account_raw,
                "AccountStd": account_std,
                "AccountOwner": owner,
                "Usage": usage,
                "ServerAddress": server_addr,
                "ManualUploadFlag": manual_flag,
                "ManualUploadFlagNorm": normalize_manual_upload_flag(manual_flag) if manual_flag_col else "",
                "ExportCheckResult": check_result,
                "IsSystemAccount": bool(is_system),
                "ExpectedStatus": expected_status,
                "SkipReason": skip_reason,
            }
        )

    unique_accounts = sorted(set(valid_accounts))
    base_summary["ExpectedCount"] = len(unique_accounts)
    base_summary["SkippedCount"] = skipped_count
    base_summary["ExpectedAccounts"] = " | ".join(unique_accounts)

    if unique_accounts:
        base_summary["ExpectedSetStatus"] = "BUILT"
        base_summary["ExpectedSetNote"] = "已从账号审阅导出构建 expected account set。"
    elif skipped_count == len(filtered):
        base_summary["ExpectedSetStatus"] = "ONLY_SKIPPED_ACCOUNTS"
        base_summary["ExpectedSetNote"] = "审阅对象存在，但全部为空账号或 SYSTEM/系统账号。"
    else:
        base_summary["ExpectedSetStatus"] = "EMPTY_AFTER_FILTER"
        base_summary["ExpectedSetNote"] = "审阅对象存在，但过滤后没有有效账号。"

    return expected_rows, base_summary


# ============================================================
# 5. Evidence plan：判断后续怎么查
# ============================================================

def decide_evidence_plan(task: pd.Series, expected_summary: Dict[str, Any]) -> Dict[str, Any]:
    """
    根据 Node、ReferTo、服务器类型、expected set 情况，判断后续证据处理方式。
    """
    node = clean_text(task.get("检查节点", ""))
    server_type_raw = clean_text(task.get("服务器类型", ""))
    server_type_norm = normalize_server_type(server_type_raw)
    refer_to = clean_text(task.get("ReferTo", ""))
    review_object = expected_summary.get("ReviewObject", "")
    expected_status = expected_summary.get("ExpectedSetStatus", "")
    expected_count = int(expected_summary.get("ExpectedCount", 0) or 0)
    manual_overall = clean_text(expected_summary.get("ManualUploadOverall", "UNKNOWN"))
    manual_status = clean_text(expected_summary.get("ManualUploadStatus", ""))
    manual_note = clean_text(expected_summary.get("ManualUploadNote", ""))

    linked_to = NODE_LINKED_TO.get(node, "")
    supporting_sheet = DEFAULT_SUPPORTING_SHEET.get(node, "")

    # 默认返回
    plan = {
        "NeedManualEvidence": "待确认",
        "EvidenceFormat": "待确认",
        "SupportingSheet": supporting_sheet,
        "EvidencePurpose": "待确认",
        "ParseMethod": "manual_review",
        "EvidencePlanStatus": "NEED_REVIEW",
        "LinkedToNode": linked_to,
        "EvidencePlanNote": "",
        "ServerTypeNormalized": server_type_norm,
    }

    def plan_from_manual_upload(
        *,
        evidence_format: str,
        evidence_purpose: str,
        parse_when_uploaded: str,
        uploaded_status: str = "READY_FOR_EVIDENCE",
        missing_status: str = "MISSING_MANUAL_EVIDENCE",
    ) -> Dict[str, Any]:
        """根据账号审阅导出里的 手工上传标识 决定证据计划。"""
        if manual_overall == "Y":
            return {
                "NeedManualEvidence": "是",
                "EvidenceFormat": evidence_format,
                "EvidencePurpose": evidence_purpose,
                "ParseMethod": parse_when_uploaded,
                "EvidencePlanStatus": uploaded_status,
                "EvidencePlanNote": manual_note,
            }
        if manual_overall == "PARTIAL":
            return {
                "NeedManualEvidence": "部分上传",
                "EvidenceFormat": evidence_format,
                "EvidencePurpose": evidence_purpose,
                "ParseMethod": parse_when_uploaded,
                "EvidencePlanStatus": "PARTIAL_MANUAL_EVIDENCE",
                "EvidencePlanNote": manual_note,
            }
        if manual_overall in {"N", "N_OR_REVIEW"}:
            return {
                "NeedManualEvidence": "否",
                "EvidenceFormat": "无手工上传材料",
                "EvidencePurpose": evidence_purpose,
                "ParseMethod": "no_manual_upload",
                "EvidencePlanStatus": missing_status if manual_overall == "N" else "NEED_REVIEW_MANUAL_UPLOAD_FLAG",
                "EvidencePlanNote": manual_note,
            }
        return {
            "NeedManualEvidence": "待确认",
            "EvidenceFormat": evidence_format,
            "EvidencePurpose": evidence_purpose,
            "ParseMethod": "manual_upload_flag_review",
            "EvidencePlanStatus": "NEED_REVIEW_MANUAL_UPLOAD_FLAG",
            "EvidencePlanNote": manual_note or f"无法根据手工上传标识判断，状态={manual_status}。",
        }

    # Refer 优先级最高：有 ReferTo 就跳 TargetSystem，不在本系统里硬找 supporting。
    if refer_to:
        plan.update(
            {
                "NeedManualEvidence": "看TargetSystem",
                "EvidenceFormat": "Refer",
                "EvidencePurpose": "引用公共系统或公共平台凭证",
                "ParseMethod": "refer_to_target",
                "EvidencePlanStatus": "REFER_TO_TARGET",
                "EvidencePlanNote": f"该任务 ReferTo={refer_to}，后续应跳转 TargetSystem 对应文件夹。",
            }
        )
        return plan

    # 没有 expected set 的通用提示
    if expected_status in {
        "ACCOUNT_EXPORT_FILE_NOT_FOUND",
        "EXPORT_REQUIRED_COLUMNS_MISSING",
        "EXPORT_MISSING_REVIEW_OBJECT_COL",
        "EXPORT_MISSING_ACCOUNT_COL",
        "EXPORT_READ_ERROR",
        "REVIEW_OBJECT_NOT_FOUND",
    }:
        # 仍然生成 evidence plan，但状态提示需要先修复导出/字段问题
        plan["EvidencePlanStatus"] = "BLOCKED_BY_EXPECTED_SET"
        plan["EvidencePlanNote"] = f"Expected set 状态为 {expected_status}，需先确认账号审阅导出。"

    # 应用层管理员：是否有手工 supporting，直接看账号审阅导出的“手工上传标识”。
    if node == "需审阅应用层管理员账号":
        plan.update(
            plan_from_manual_upload(
                evidence_format="Excel中截图 / 截图 / Excel",
                evidence_purpose="核对应用层管理员账号是否有 supporting 覆盖",
                parse_when_uploaded="account_presence_or_matrix_review",
                uploaded_status="READY_FOR_EVIDENCE" if expected_count > 0 else "READY_FOR_EVIDENCE_BUT_EXPECTED_SET_REVIEW",
            )
        )
        return plan

    # 上传应用层账号截图：不是账号比对节点。
    # 它只看关联审阅对象（应用层）的“手工上传标识”：Y=已上传，N=未上传。
    if node == "需上传应用层账号截图":
        plan.update(
            plan_from_manual_upload(
                evidence_format="Excel中截图 / 截图",
                evidence_purpose="检查应用层账号 supporting 是否存在，并支撑应用层管理员账号审阅",
                parse_when_uploaded="evidence_exists_from_manual_upload_flag",
                uploaded_status="EVIDENCE_UPLOADED_BY_MANUAL_FLAG",
                missing_status="MISSING_EVIDENCE_BY_MANUAL_FLAG",
            )
        )
        return plan

    # 发版工具前台：是否上传 supporting，直接看“手工上传标识”。
    if node == "需审阅发版工具发版账号":
        plan.update(
            plan_from_manual_upload(
                evidence_format="Excel中截图 / 截图 / Excel",
                evidence_purpose="核对发布账号、发布创建人、流水线执行者等是否覆盖",
                parse_when_uploaded="account_presence_or_matrix_review",
                uploaded_status="READY_FOR_EVIDENCE" if expected_count > 0 else "READY_FOR_EVIDENCE_BUT_EXPECTED_SET_REVIEW",
            )
        )
        return plan

    # 数据库账号：通常不需要手工截图，先检查系统导出可用。
    if node == "需审阅数据库账号":
        if expected_count > 0:
            status = "EXPORT_AVAILABLE"
            note = "数据库账号通常由系统抓取；第一版先以账号审阅导出存在有效数据库账号作为通过导出可用性检查。"
        elif expected_status == "ONLY_SKIPPED_ACCOUNTS":
            status = "ONLY_SKIPPED_ACCOUNTS"
            note = "数据库账号审阅对象存在，但过滤后只剩 SYSTEM/空账号，暂不进入 OCR。"
        else:
            status = "REVIEW_EXPORT_EMPTY"
            note = "Scope 要查数据库账号，但账号审阅导出未筛出有效数据库账号。"

        plan.update(
            {
                "NeedManualEvidence": "否",
                "EvidenceFormat": "系统导出",
                "EvidencePurpose": "确认数据库账号导出存在有效账号",
                "ParseMethod": "export_available",
                "EvidencePlanStatus": status,
                "EvidencePlanNote": note,
            }
        )
        return plan

    # 服务器类：不再靠复杂服务器规则判断是否上传截图。
    # 是否手工上传材料，以账号审阅导出里的“手工上传标识”为准；服务器类型只作为后续解释/复核字段。
    if node in {"需审阅应用层服务器账号", "需审阅数据库服务器账号", "需审阅发版工具后台账号"}:
        plan.update(
            plan_from_manual_upload(
                evidence_format="Excel中截图 / Excel / 系统导出",
                evidence_purpose=f"服务器类节点；服务器类型={server_type_norm}，是否上传材料以手工上传标识为准",
                parse_when_uploaded="server_evidence_review_from_manual_upload",
                uploaded_status="READY_FOR_SERVER_EVIDENCE_REVIEW" if expected_count > 0 else "READY_FOR_SERVER_EVIDENCE_BUT_EXPECTED_SET_REVIEW",
                missing_status="NO_MANUAL_UPLOAD_FLAG",
            )
        )
        return plan

    # 开发人员清单：来自外部 supporting，不一定来自账号审阅导出。
    if node == "需上传开发人员清单":
        plan.update(
            {
                "NeedManualEvidence": "是",
                "EvidenceFormat": "Excel / TXT / Excel中截图",
                "EvidencePurpose": "确认开发人员清单是否上传且可读取",
                "ParseMethod": "evidence_exists",
                "EvidencePlanStatus": "READY_FOR_EVIDENCE_EXISTENCE_CHECK",
                "EvidencePlanNote": "开发人员清单不一定来自账号审阅导出；第一版先检查文件存在和可读性。",
            }
        )
        return plan

    # 派生检查：开发人员权限 / SOD
    if node in {"开发人员在生产环境无只读以上权限", "无前后台管理员SOD问题"}:
        plan.update(
            {
                "NeedManualEvidence": "派生/另行判断",
                "EvidenceFormat": "Derived",
                "EvidencePurpose": "基于多个账号集合做交集或权限等级判断",
                "ParseMethod": "derived_intersection",
                "EvidencePlanStatus": "PENDING_DERIVED_CHECK",
                "EvidencePlanNote": "该节点不依赖单张截图；后续需要开发人员 set、应用层/DB/server 管理员 set、权限等级等集合。",
            }
        )
        return plan

    return plan


# ============================================================
# 6. Compare result 占位
# ============================================================

def build_compare_placeholder(task: pd.Series, expected_summary: Dict[str, Any], evidence_plan: Dict[str, Any]) -> Dict[str, Any]:
    node = clean_text(task.get("检查节点", ""))
    parse_method = evidence_plan.get("ParseMethod", "")
    expected_status = expected_summary.get("ExpectedSetStatus", "")
    expected_count = int(expected_summary.get("ExpectedCount", 0) or 0)

    result = {
        "TaskID": expected_summary.get("TaskID", ""),
        "System": expected_summary.get("System", ""),
        "TargetSystem": expected_summary.get("TargetSystem", ""),
        "Node": node,
        "ReviewObject": expected_summary.get("ReviewObject", ""),
        "ParseMethod": parse_method,
        "ExpectedAccounts": expected_summary.get("ExpectedAccounts", ""),
        "ExpectedCount": expected_count,
        "ActualAccounts": "",
        "ActualCount": "",
        "MissingAccounts": "",
        "ExtraAccounts": "",
        "ExpectedUsage": "",
        "ActualUsage": "",
        "UsageMatchStatus": "PENDING",
        "CheckResult": "",
        "Remark": "",
    }

    if parse_method == "export_available":
        if expected_count > 0:
            result["CheckResult"] = "PASS_EXPORT_AVAILABLE"
            result["Remark"] = "账号审阅导出中存在有效账号；该节点第一版不要求 OCR。"
        elif expected_status == "ONLY_SKIPPED_ACCOUNTS":
            result["CheckResult"] = "SKIPPED_ONLY_SYSTEM_ACCOUNTS"
            result["Remark"] = "审阅对象存在，但过滤后仅剩 SYSTEM/系统账号或空账号。"
        else:
            result["CheckResult"] = "REVIEW_EXPORT_EMPTY"
            result["Remark"] = "账号审阅导出未筛出有效账号，需要人工确认。"
        return result

    if parse_method in {"account_presence_or_matrix_review"}:
        if expected_count > 0:
            result["CheckResult"] = "PENDING_OCR_OR_MATRIX_REVIEW"
            result["Remark"] = "待 OCR / 矩阵解析后进行账号 presence 比对；用途/权限勾选暂不自动判断。"
        elif expected_status == "ONLY_SKIPPED_ACCOUNTS":
            result["CheckResult"] = "SKIPPED_ONLY_SYSTEM_ACCOUNTS"
            result["Remark"] = "过滤后无有效账号，暂不进入 OCR。"
        else:
            result["CheckResult"] = "REVIEW_EXPECTED_SET"
            result["Remark"] = "expected set 未构建成功，需先确认账号审阅导出。"
        return result

    if parse_method in {"evidence_exists_then_link_to_app_admin", "evidence_exists_from_manual_upload_flag"}:
        plan_status = evidence_plan.get("EvidencePlanStatus", "")
        if plan_status == "EVIDENCE_UPLOADED_BY_MANUAL_FLAG":
            result["CheckResult"] = "EVIDENCE_FOUND_BY_MANUAL_FLAG"
            result["Remark"] = "手工上传标识为 Y，视为已上传对应截图/材料；后续如需可再定位 supporting 文件。"
        elif plan_status == "MISSING_EVIDENCE_BY_MANUAL_FLAG":
            result["CheckResult"] = "MISSING_EVIDENCE"
            result["Remark"] = "手工上传标识为 N，表示未上传对应截图/材料。"
        elif plan_status == "PARTIAL_MANUAL_EVIDENCE":
            result["CheckResult"] = "REVIEW_PARTIAL_EVIDENCE"
            result["Remark"] = "手工上传标识存在部分 Y / 部分 N，需要人工复核。"
        else:
            result["CheckResult"] = "REVIEW_MANUAL_UPLOAD_FLAG"
            result["Remark"] = evidence_plan.get("EvidencePlanNote", "需确认手工上传标识。")
        return result

    if parse_method == "no_manual_upload":
        plan_status = evidence_plan.get("EvidencePlanStatus", "")
        if plan_status == "MISSING_MANUAL_EVIDENCE":
            result["CheckResult"] = "MISSING_EVIDENCE"
            result["Remark"] = "手工上传标识为 N，未上传材料。"
        else:
            result["CheckResult"] = "NO_MANUAL_UPLOAD"
            result["Remark"] = evidence_plan.get("EvidencePlanNote", "手工上传标识显示未上传或需复核。")
        return result

    if parse_method == "refer_to_target":
        result["CheckResult"] = "PENDING_REFER_TARGET_CHECK"
        result["Remark"] = "该节点需要跳转 TargetSystem / ReferTo 对应凭证。"
        return result

    if parse_method in {"windows_group_check", "mixed_server_type_review", "server_evidence_review_from_manual_upload"}:
        result["CheckResult"] = "PENDING_SERVER_EVIDENCE_REVIEW"
        result["Remark"] = "手工上传标识为 Y 或部分上传，后续需读取 supporting/截图/Excel；服务器类型和用户组规则留作复核。"
        return result

    if parse_method in {"export_available_or_pam_review"}:
        result["CheckResult"] = "PENDING_EXPORT_OR_PAM_REVIEW"
        result["Remark"] = "服务器类 Linux/PAM 节点第一版先保守标记为待复核。"
        return result

    if parse_method == "evidence_exists":
        result["CheckResult"] = "PENDING_EVIDENCE_EXISTENCE_CHECK"
        result["Remark"] = "待检查开发人员清单等 supporting 是否存在。"
        return result

    if parse_method == "derived_intersection":
        result["CheckResult"] = "PENDING_DERIVED_CHECK"
        result["Remark"] = "待后续基于多个账号集合做交集/权限等级判断。"
        return result

    result["CheckResult"] = "NEED_REVIEW"
    result["Remark"] = evidence_plan.get("EvidencePlanNote", "未配置明确比对方式。")
    return result


# ============================================================
# 7. 主流程：生成 execution plan workbook
# ============================================================

def load_todo_with_folder(path: str | Path, sheet_name: str = "todo_with_folder") -> pd.DataFrame:
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"找不到 todo_with_folder 文件：{path}")

    # 优先读取指定 sheet；如果没有，就读取第一个 sheet
    try:
        df = pd.read_excel(path, sheet_name=sheet_name)
    except ValueError:
        df = pd.read_excel(path, sheet_name=0)

    required = ["TaskID", "System", "检查节点"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"todo_with_folder 缺少必要列：{missing}")

    if "TargetSystem" not in df.columns:
        df["TargetSystem"] = df["System"]

    if "ResolvedFolderPath" not in df.columns:
        df["ResolvedFolderPath"] = ""

    if "服务器类型" not in df.columns:
        df["服务器类型"] = ""

    if "ReferTo" not in df.columns:
        df["ReferTo"] = ""

    return df


def build_execution_plan(
    todo_with_folder_path: str | Path,
    output_path: str | Path = "output/execution_plan.xlsx",
    export_sheet_name: Optional[str] = None,
) -> Dict[str, str]:
    todo_df = load_todo_with_folder(todo_with_folder_path)

    expected_account_rows: List[Dict[str, Any]] = []
    evidence_plan_rows: List[Dict[str, Any]] = []
    compare_result_rows: List[Dict[str, Any]] = []

    # 缓存每个系统文件夹读取出来的账号审阅导出，避免重复读
    export_cache: Dict[str, AccountExportReadResult] = {}

    for _, task in todo_df.iterrows():
        node = clean_text(task.get("检查节点", ""))
        review_object = NODE_TO_REVIEW_OBJECT.get(node, "")
        folder_path = clean_text(task.get("ResolvedFolderPath", ""))

        # 读取账号审阅导出
        if not folder_path:
            export_result = AccountExportReadResult(
                status="NO_RESOLVED_FOLDER_PATH",
                message="该任务没有 ResolvedFolderPath，无法定位账号审阅导出。",
            )
        else:
            cache_key = folder_path
            if cache_key not in export_cache:
                export_file = find_account_export_file(folder_path)
                if export_file is None:
                    export_cache[cache_key] = AccountExportReadResult(
                        status="ACCOUNT_EXPORT_FILE_NOT_FOUND",
                        message=f"文件夹中未找到文件名包含“账号审阅导出”的 Excel：{folder_path}",
                    )
                else:
                    export_cache[cache_key] = read_account_export(export_file, sheet_name=export_sheet_name)

            export_result = export_cache[cache_key]

        # Evidence-only node（例如：需上传应用层账号截图）不生成账号 expected set。
        # 但它需要读取关联审阅对象（应用层）的“手工上传标识”，判断截图/材料是否上传。
        if node in EVIDENCE_ONLY_NODE_REVIEW_OBJECT:
            linked_review_object = EVIDENCE_ONLY_NODE_REVIEW_OBJECT[node]
            _, linked_summary = build_expected_accounts_for_task(export_result, task, linked_review_object)
            summary = dict(linked_summary)
            summary["ReviewObject"] = ""
            summary["LinkedReviewObject"] = linked_review_object
            summary["ExpectedCount"] = 0
            summary["ExpectedAccounts"] = ""
            summary["ExpectedSetStatus"] = "NOT_ACCOUNT_CHECK_NODE"
            summary["ExpectedSetNote"] = (
                f"该节点不生成账号 expected set；只读取关联审阅对象={linked_review_object} 的手工上传标识判断材料是否上传。"
            )
            rows = []
        else:
            rows, summary = build_expected_accounts_for_task(export_result, task, review_object)
            expected_account_rows.extend(rows)

        plan = decide_evidence_plan(task, summary)

        evidence_plan_rows.append(
            {
                "TaskID": summary.get("TaskID", ""),
                "System": summary.get("System", ""),
                "TargetSystem": summary.get("TargetSystem", ""),
                "Node": summary.get("Node", ""),
                "ReviewObject": summary.get("ReviewObject", ""),
                "ServerType": summary.get("ServerType", ""),
                "ServerTypeNormalized": plan.get("ServerTypeNormalized", ""),
                "ReferTo": summary.get("ReferTo", ""),
                "ScopeLabel": summary.get("ScopeLabel", ""),
                "ExportFile": summary.get("ExportFile", ""),
                "ExportSheet": summary.get("ExportSheet", ""),
                "RawMatchedRows": summary.get("RawMatchedRows", 0),
                "ExpectedCount": summary.get("ExpectedCount", 0),
                "SkippedCount": summary.get("SkippedCount", 0),
                "ExpectedAccounts": summary.get("ExpectedAccounts", ""),
                "ExpectedSetStatus": summary.get("ExpectedSetStatus", ""),
                "ExpectedSetNote": summary.get("ExpectedSetNote", ""),
                "ManualUploadColumn": summary.get("ManualUploadColumn", ""),
                "ManualUploadYCount": summary.get("ManualUploadYCount", 0),
                "ManualUploadNCount": summary.get("ManualUploadNCount", 0),
                "ManualUploadBlankCount": summary.get("ManualUploadBlankCount", 0),
                "ManualUploadOtherValues": summary.get("ManualUploadOtherValues", ""),
                "ManualUploadOverall": summary.get("ManualUploadOverall", ""),
                "ManualUploadStatus": summary.get("ManualUploadStatus", ""),
                "ManualUploadNote": summary.get("ManualUploadNote", ""),
                "NeedManualEvidence": plan.get("NeedManualEvidence", ""),
                "EvidenceFormat": plan.get("EvidenceFormat", ""),
                "SupportingSheet": plan.get("SupportingSheet", ""),
                "EvidencePurpose": plan.get("EvidencePurpose", ""),
                "ParseMethod": plan.get("ParseMethod", ""),
                "LinkedToNode": plan.get("LinkedToNode", ""),
                "EvidencePlanStatus": plan.get("EvidencePlanStatus", ""),
                "EvidencePlanNote": plan.get("EvidencePlanNote", ""),
            }
        )

        # compare_result 只保留手工上传标识为 Y 的任务
        # （没有手工上传材料的节点不需要 OCR 比对，没必要占 compare_result 行）
        manual_overall = summary.get("ManualUploadOverall", "")
        if manual_overall == "Y":
            compare_result_rows.append(build_compare_placeholder(task, summary, plan))

    expected_df = pd.DataFrame(expected_account_rows)
    evidence_plan_df = pd.DataFrame(evidence_plan_rows)
    compare_result_df = pd.DataFrame(compare_result_rows)

    # 即使为空，也输出固定列，避免后续读取报错。
    expected_columns = [
        "TaskID",
        "System",
        "TargetSystem",
        "Node",
        "ReviewObject",
        "ExportFile",
        "ExportSheet",
        "SourceRowIndex",
        "ReviewObjectValue",
        "AccountRaw",
        "AccountStd",
        "AccountOwner",
        "Usage",
        "ServerAddress",
        "ManualUploadFlag",
        "ManualUploadFlagNorm",
        "ExportCheckResult",
        "IsSystemAccount",
        "ExpectedStatus",
        "SkipReason",
    ]
    for col in expected_columns:
        if col not in expected_df.columns:
            expected_df[col] = ""
    expected_df = expected_df[expected_columns]

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        expected_df.to_excel(writer, sheet_name="expected_accounts", index=False)
        evidence_plan_df.to_excel(writer, sheet_name="evidence_plan", index=False)
        compare_result_df.to_excel(writer, sheet_name="compare_result", index=False)

    return {
        "execution_plan": str(output_path),
    }


# ============================================================
# 8. CLI
# ============================================================

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Build execution plan from todo_list_with_folder and account export files."
    )
    parser.add_argument(
        "--todo-with-folder",
        required=True,
        help="System-Folder Mapper 输出的 todo_list_with_folder.xlsx。",
    )
    parser.add_argument(
        "--output",
        default="output/execution_plan.xlsx",
        help="输出 execution_plan.xlsx 的路径。",
    )
    parser.add_argument(
        "--export-sheet",
        default=None,
        help="账号审阅导出 Excel 的 sheet 名。不填则自动扫描包含 审阅对象 + 账号名 的 sheet。",
    )

    args = parser.parse_args()

    files = build_execution_plan(
        todo_with_folder_path=args.todo_with_folder,
        output_path=args.output,
        export_sheet_name=args.export_sheet,
    )

    print(json.dumps(files, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
