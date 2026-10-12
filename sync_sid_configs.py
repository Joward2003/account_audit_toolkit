"""用权威 System/SID 表批量同步项目内所有系统级配置。"""

from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Iterable

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from sid_utils import normalize_sid
from sync_application_admin_role_config import (
    build_rows as build_role_rows,
    write_config as write_role_config,
)


APP_DIR = Path(__file__).resolve().parent
DEFAULT_SOURCE = APP_DIR.parent / "各系统应用层管理员角色定义范围.xlsx"
DEFAULT_LEGACY_DIR = APP_DIR.parent / "account_audit_toolkit" / "configs"

FOLDER_HEADERS = ["SID", "实际文件夹名", "备注"]
RULE_HEADERS = [
    "SID",
    "是否合规支持平台",
    "是否PMS_VMS",
    "数据库服务器结论",
    "数据库服务器原因",
    "数据库账号结论",
    "数据库账号原因",
    "是否检查SOD",
    "SOD原因",
    "发版工具前台结论",
    "发版工具前台原因",
    "发版工具后台结论",
    "发版工具后台原因",
    "警告提示",
    "备注",
    "来源",
    "确认人",
    "确认日期",
    "是否启用",
]
MANUAL_HEADERS = [
    "SID",
    "检查节点",
    "覆盖结论",
    "是否仍需证据",
    "引用SID",
    "原因",
    "是否启用",
]

RULE_COLUMN_MAP = {
    "db_server_label": "数据库服务器结论",
    "db_server_reason": "数据库服务器原因",
    "db_account_label": "数据库账号结论",
    "db_account_reason": "数据库账号原因",
    "sod_required": "是否检查SOD",
    "sod_reason": "SOD原因",
    "release_front_label": "发版工具前台结论",
    "release_front_reason": "发版工具前台原因",
    "release_back_label": "发版工具后台结论",
    "release_back_reason": "发版工具后台原因",
    "warning": "警告提示",
}


def normalize_system_name(value: object) -> str:
    """仅忽略大小写和常见连接符；不做相似度猜测。"""
    text = "" if value is None else str(value).strip().upper()
    return re.sub(r"[\s_\-/\\（）()&]+", "", text)


def load_sid_source(source_path: Path) -> tuple[pd.DataFrame, dict[str, dict[str, str]]]:
    df = pd.read_excel(source_path, dtype=str, keep_default_na=False)
    missing = {"System", "SID"} - set(df.columns)
    if missing:
        raise ValueError(f"{source_path} 缺少字段：{sorted(missing)}")

    records: list[dict[str, str]] = []
    by_name: dict[str, dict[str, str]] = {}
    by_sid: dict[str, str] = {}
    for _, row in df.iterrows():
        system = str(row["System"]).strip()
        sid = normalize_sid(row["SID"])
        if not system or not sid:
            continue
        name_key = normalize_system_name(system)
        if sid in by_sid and by_sid[sid] != system:
            raise ValueError(f"来源表 SID 冲突：{sid} 同时对应 {by_sid[sid]} / {system}")
        if name_key in by_name and by_name[name_key]["sid"] != sid:
            raise ValueError(
                f"来源表系统名称冲突：{system} 同时对应 {by_name[name_key]['sid']} / {sid}"
            )
        record = {"system": system, "sid": sid}
        by_sid[sid] = system
        by_name[name_key] = record
        records.append(record)
    return pd.DataFrame(records), by_name


def write_workbook(
    target_path: Path,
    sheet_name: str,
    headers: list[str],
    rows: Iterable[Iterable[object]],
) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.append(headers)
    for row in rows:
        ws.append(list(row))

    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    for cell in ws["A"]:
        cell.number_format = "@"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{ws.cell(1, len(headers)).column_letter}{max(ws.max_row, 1)}"
    for index, header in enumerate(headers, start=1):
        width = 16 if header == "SID" else min(max(len(header) * 2 + 4, 16), 48)
        ws.column_dimensions[ws.cell(1, index).column_letter].width = width
    target_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(target_path)


def build_folder_rows(
    sid_df: pd.DataFrame,
    by_name: dict[str, dict[str, str]],
    legacy_dir: Path,
) -> tuple[list[list[str]], list[dict[str, str]]]:
    legacy_path = legacy_dir / "system_folder_mapping.xlsx"
    legacy_by_sid: dict[str, dict[str, str]] = {}
    unresolved_rows: list[list[str]] = []
    pending: list[dict[str, str]] = []
    if legacy_path.exists():
        legacy_df = pd.read_excel(legacy_path, dtype=str, keep_default_na=False)
        for _, row in legacy_df.iterrows():
            name = str(row.get("system", "")).strip()
            source = by_name.get(normalize_system_name(name))
            if not source:
                pending.append({
                    "配置表": legacy_path.name,
                    "原系统名称": name,
                    "待补原因": "权威来源表中找不到对应 SID",
                })
                unresolved_rows.append(["", name, "暂无SID"])
                continue
            legacy_by_sid[source["sid"]] = {
                "actual": str(row.get("actual_folder_name", "")).strip(),
                "note": str(row.get("note", "")).strip(),
            }

    rows: list[list[str]] = []
    for _, record in sid_df.iterrows():
        sid, system = record["sid"], record["system"]
        legacy = legacy_by_sid.get(sid)
        if legacy:
            actual = legacy["actual"] or system
            note = f"由旧映射迁移；{legacy['note']}".rstrip("；")
        else:
            actual = system
            note = "按权威 SID 表初始化；如 supporting 实际目录不同，请只修改“实际文件夹名”"
        rows.append([sid, actual, note])
    rows.extend(unresolved_rows)
    return rows, pending


def build_rule_rows(
    by_name: dict[str, dict[str, str]],
    legacy_dir: Path,
) -> tuple[list[list[str]], list[dict[str, str]]]:
    legacy_path = legacy_dir / "system_rule_overrides.csv"
    if not legacy_path.exists():
        return [], []
    legacy_df = pd.read_csv(legacy_path, dtype=str, keep_default_na=False)
    rows: list[list[str]] = []
    pending: list[dict[str, str]] = []
    seen_sids: set[str] = set()
    for _, legacy in legacy_df.iterrows():
        name = str(legacy.get("system", "")).strip()
        source = by_name.get(normalize_system_name(name))
        if not source:
            pending.append({
                "配置表": legacy_path.name,
                "原系统名称": name,
                "待补原因": "权威来源表中找不到对应 SID",
            })
            values = {
                target: str(legacy.get(source_col, "")).strip()
                for source_col, target in RULE_COLUMN_MAP.items()
            }
            values.update({
                "SID": "",
                "备注": "暂无SID",
                "来源": f"迁移自 {legacy_path.name}（原系统名：{name}）",
                "确认人": "",
                "确认日期": "",
                "是否启用": "是",
            })
            rows.append([values.get(header, "") for header in RULE_HEADERS])
            continue
        sid = source["sid"]
        if sid in seen_sids:
            raise ValueError(f"{legacy_path} 多行规则迁移到同一 SID：{sid}")
        seen_sids.add(sid)
        values = {target: str(legacy.get(source_col, "")).strip()
                  for source_col, target in RULE_COLUMN_MAP.items()}
        values.update({
            "SID": sid,
            "备注": "",
            "来源": f"迁移自 {legacy_path.name}（原系统名：{name}）",
            "确认人": "",
            "确认日期": "",
            "是否启用": "是",
        })
        rows.append([values.get(header, "") for header in RULE_HEADERS])
    return rows, pending


def build_manual_rows(
    by_name: dict[str, dict[str, str]],
    legacy_dir: Path,
) -> tuple[list[list[str]], list[dict[str, str]]]:
    legacy_path = legacy_dir / "manual_override.xlsx"
    if not legacy_path.exists():
        return [], []
    legacy_df = pd.read_excel(legacy_path, dtype=str, keep_default_na=False)
    rows: list[list[str]] = []
    pending: list[dict[str, str]] = []
    for _, legacy in legacy_df.iterrows():
        name = str(legacy.get("system", "")).strip()
        source = by_name.get(normalize_system_name(name))
        if not source:
            pending.append({
                "配置表": legacy_path.name,
                "原系统名称": name,
                "待补原因": "权威来源表中找不到对应 SID",
            })
            continue
        refer_name = str(legacy.get("refer_to", "")).strip()
        refer_source = by_name.get(normalize_system_name(refer_name)) if refer_name else None
        if refer_name and not refer_source:
            pending.append({
                "配置表": legacy_path.name,
                "原系统名称": refer_name,
                "待补原因": f"{name} 的引用目标缺少 SID",
            })
        values = {
            "SID": source["sid"],
            "检查节点": str(legacy.get("node", "")).strip(),
            "覆盖结论": str(legacy.get("override_label", "")).strip(),
            "是否仍需证据": str(legacy.get("requires_evidence", "")).strip(),
            "引用SID": refer_source["sid"] if refer_source else "",
            "原因": str(legacy.get("reason", "")).strip(),
            "是否启用": str(legacy.get("enabled", "")).strip(),
        }
        rows.append([values.get(header, "") for header in MANUAL_HEADERS])
    return rows, pending


def main() -> None:
    parser = argparse.ArgumentParser(description="同步所有 SID 系统级配置")
    parser.add_argument("--source", default=str(DEFAULT_SOURCE), help="权威 System/SID 来源表")
    parser.add_argument("--config-dir", default=str(APP_DIR / "configs"))
    parser.add_argument("--legacy-dir", default=str(DEFAULT_LEGACY_DIR))
    parser.add_argument("--report", default=str(APP_DIR / "output" / "SID配置同步报告.xlsx"))
    args = parser.parse_args()

    source_path = Path(args.source)
    config_dir = Path(args.config_dir)
    legacy_dir = Path(args.legacy_dir)
    sid_df, by_name = load_sid_source(source_path)

    folder_rows, folder_pending = build_folder_rows(sid_df, by_name, legacy_dir)
    rule_rows, rule_pending = build_rule_rows(by_name, legacy_dir)
    manual_rows, manual_pending = build_manual_rows(by_name, legacy_dir)
    role_rows, skipped_role_sids = build_role_rows(source_path)

    write_workbook(
        config_dir / "系统文件夹映射表.xlsx",
        "系统文件夹映射",
        FOLDER_HEADERS,
        folder_rows,
    )
    write_workbook(
        config_dir / "系统规则特殊覆盖表.xlsx",
        "系统规则特殊覆盖",
        RULE_HEADERS,
        rule_rows,
    )
    write_workbook(
        config_dir / "人工覆盖表（预留）.xlsx",
        "人工覆盖",
        MANUAL_HEADERS,
        manual_rows,
    )
    write_role_config(role_rows, config_dir / "应用层管理员角色配置表.xlsx")

    pending = folder_pending + rule_pending + manual_pending
    report_rows = [
        ["已同步", "系统文件夹映射表.xlsx", len(folder_rows), "非空 SID 唯一；空 SID 为待补"],
        ["已同步", "系统规则特殊覆盖表.xlsx", len(rule_rows), "非空 SID 唯一；空 SID 为待补"],
        ["已同步", "人工覆盖表（预留）.xlsx", len(manual_rows), "SID 唯一"],
        ["已同步", "应用层管理员角色配置表.xlsx", len(role_rows), "SID+角色关键词唯一"],
    ]
    report_path = Path(args.report)
    write_workbook(
        report_path,
        "同步汇总",
        ["状态", "配置表", "写入行数", "唯一键"],
        report_rows,
    )
    if pending or skipped_role_sids:
        with pd.ExcelWriter(
            report_path, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            pd.DataFrame(
                pending,
                columns=["配置表", "原系统名称", "待补原因"],
            ).to_excel(writer, sheet_name="待补SID", index=False)
            pd.DataFrame(
                {"SID": skipped_role_sids, "原因": "来源角色为空/N/A/TBD"}
            ).to_excel(writer, sheet_name="角色待维护", index=False)

    print(
        f"同步完成：文件夹 {len(folder_rows)} 行，特殊规则 {len(rule_rows)} 行，"
        f"人工覆盖 {len(manual_rows)} 行，管理员角色 {len(role_rows)} 行。"
    )
    if pending:
        print("以下旧系统缺少权威 SID，未猜填：" + " | ".join(
            sorted({item["原系统名称"] for item in pending})
        ))
    print(f"同步报告：{report_path}")


if __name__ == "__main__":
    main()
