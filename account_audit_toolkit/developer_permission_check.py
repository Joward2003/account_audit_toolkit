#!/usr/bin/env python3
"""
开发人员权限检查 & 管理员重合分析脚本
============================================

功能:
1. 读取开发人员名单和账号审阅导出
2. 判断开发人员的域账号/姓名在账号审阅中的权限（用途列）是否为「只读」
3. 检查「应用层」（审阅对象=应用层）和「数据库/全部服务器」
   （审阅对象=数据库/应用服务器/发布工具服务器/数据库服务器）中
   同时具有管理员/读写权限的人员是否存在重合，输出 Excel 标注。

用法:
  python developer_permission_check.py

输出:
  output/developer_permission_check.xlsx
    - Sheet "权限检查": 开发人员权限检查明细
    - Sheet "管理员重合": 应用层 & 数据库/服务器管理员重合
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

# ---------------------------------------------------------------------------
# 配置
# ---------------------------------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent
SUPPORTING_DIR = BASE_DIR / "supporting"
OUTPUT_DIR = BASE_DIR / "account_audit_toolkit" / "output"
OUTPUT_FILE = OUTPUT_DIR / "developer_permission_check.xlsx"

# 系统文件夹列表
SYSTEMS = ["OPM", "YumcZoo"]

# 审阅对象分类
# 应用层：仅「应用层」
APP_LAYER_OBJECTS = {"应用层"}
# 其他（数据库/服务器/全部服务器）：数据库 + 各类服务器
DB_SERVER_OBJECTS = {"数据库", "应用服务器", "发布工具服务器", "数据库服务器"}

# 管理员用途
ADMIN_ROLES = {"管理员", "读写"}


# ---------------------------------------------------------------------------
# 辅助函数
# ---------------------------------------------------------------------------

def read_developer_list(system: str) -> pd.DataFrame:
    """读取开发人员清单，返回标准化 DataFrame（列: 系统, 姓名, 域账号）。"""
    path = SUPPORTING_DIR / system / "开发人员清单.xlsx"

    if system == "OPM":
        df = pd.read_excel(path, sheet_name="开发人员名单")
        df = df.rename(columns={
            "姓名（中文或英文全名）": "姓名",
        })
        df["系统"] = system
        df["域账号"] = ""  # OPM 无域账号列
        return df[["系统", "姓名", "域账号"]]
    else:
        df = pd.read_excel(path, sheet_name="Sheet1")
        df["系统"] = system
        return df[["系统", "姓名", "域账号"]]


def read_audit_export(system: str) -> pd.DataFrame:
    """读取账号审阅导出。"""
    path = SUPPORTING_DIR / system / "账号审阅导出.xlsx"
    return pd.read_excel(path, sheet_name="账号审阅数据")


def match_developer_in_audit(
    dev_row: pd.Series,
    audit_df: pd.DataFrame,
) -> pd.DataFrame:
    """
    在审阅数据中匹配一个开发人员。
    匹配策略：
    - 域账号 → 账号名（精确匹配，忽略大小写和空格）
    - 姓名 → 账号使用人（模糊包含匹配）
    """
    name = str(dev_row.get("姓名", "")).strip()
    account = str(dev_row.get("域账号", "")).strip()

    matches = pd.DataFrame()

    # 按域账号精确匹配
    if account:
        mask_account = audit_df["账号名"].astype(str).str.strip().str.upper() == account.upper()
        matches = pd.concat([matches, audit_df[mask_account]])

    # 按姓名模糊匹配（且未被域账号覆盖）
    if name:
        mask_name = audit_df["账号使用人"].astype(str).str.contains(name, na=False)
        name_matches = audit_df[mask_name]
        if not matches.empty:
            existing_idx = matches.index
            name_matches = name_matches[~name_matches.index.isin(existing_idx)]
        matches = pd.concat([matches, name_matches])

    return matches.drop_duplicates()


def check_dev_permissions(
    system: str,
    dev_df: pd.DataFrame,
    audit_df: pd.DataFrame,
) -> pd.DataFrame:
    """检查每个开发人员权限，返回结果 DataFrame。"""
    rows: List[Dict] = []

    for _, dev_row in dev_df.iterrows():
        name = str(dev_row.get("姓名", "")).strip()
        account = str(dev_row.get("域账号", "")).strip()

        matched = match_developer_in_audit(dev_row, audit_df)

        if matched.empty:
            rows.append({
                "系统": system,
                "开发人员姓名": name,
                "域账号": account,
                "账号名(审阅)": "",
                "账号使用人(审阅)": "",
                "审阅对象": "",
                "用途": "",
                "权限是否只读": "未找到匹配记录",
                "问题说明": "该开发人员在账号审阅导出中未找到对应记录",
            })
        else:
            for _, m in matched.iterrows():
                usage = str(m.get("用途", "")).strip()
                is_readonly = "是" if usage == "只读" else "否"
                problem = ""
                if usage in ADMIN_ROLES:
                    problem = f"权限为「{usage}」，非只读！"
                elif usage == "系统账号":
                    problem = "系统账号，非个人账号"
                elif usage != "只读":
                    problem = f"未知用途类型「{usage}」"

                rows.append({
                    "系统": system,
                    "开发人员姓名": name,
                    "域账号": account,
                    "账号名(审阅)": m.get("账号名", ""),
                    "账号使用人(审阅)": m.get("账号使用人", ""),
                    "审阅对象": m.get("审阅对象", ""),
                    "用途": usage,
                    "权限是否只读": is_readonly,
                    "问题说明": problem,
                })

    return pd.DataFrame(rows)


def find_admin_overlap(
    system: str,
    audit_df: pd.DataFrame,
) -> pd.DataFrame:
    """
    找出在「应用层」（审阅对象=应用层）和「数据库/全部服务器」
    （审阅对象=数据库/应用服务器/发布工具服务器/数据库服务器）
    同时具有管理员/读写权限的人员。
    返回: 系统, 重合人, 重合账号, 具体哪个重合
    """
    # 应用层管理员记录
    app_admin = audit_df[
        (audit_df["审阅对象"].isin(APP_LAYER_OBJECTS)) &
        (audit_df["用途"].isin(ADMIN_ROLES))
    ].copy()

    # 数据库/服务器管理员记录
    db_admin = audit_df[
        (audit_df["审阅对象"].isin(DB_SERVER_OBJECTS)) &
        (audit_df["用途"].isin(ADMIN_ROLES))
    ].copy()

    rows: List[Dict] = []

    # 按「账号使用人」匹配重合
    app_users = set(app_admin["账号使用人"].dropna().astype(str).str.strip())
    db_users = set(db_admin["账号使用人"].dropna().astype(str).str.strip())
    common_users = app_users & db_users

    # 过滤掉系统账号/通用账号/泛化描述
    skip_users = {
        "系统", "系统账号", "PAM", "",
        "登录服务器后的读写账号", "登录服务器后的只读账号",
        "系统默认账号",
    }

    for user in sorted(common_users):
        if user in skip_users:
            continue

        app_records = app_admin[app_admin["账号使用人"].astype(str).str.strip() == user]
        db_records = db_admin[db_admin["账号使用人"].astype(str).str.strip() == user]

        app_accounts = set(app_records["账号名"].astype(str).str.strip())
        db_accounts = set(db_records["账号名"].astype(str).str.strip())

        app_objects = set(app_records["审阅对象"].astype(str).str.strip())
        db_objects = set(db_records["审阅对象"].astype(str).str.strip())

        detail = (
            f"应用层: {'/'.join(sorted(app_objects))} "
            f"(账号: {'/'.join(sorted(app_accounts))}); "
            f"数据库/服务器: {'/'.join(sorted(db_objects))} "
            f"(账号: {'/'.join(sorted(db_accounts))})"
        )

        rows.append({
            "系统": system,
            "重合人": user,
            "重合账号": " / ".join(sorted(app_accounts | db_accounts)),
            "具体哪个重合": detail,
        })

    # 也按「账号名」匹配（同一账号可能不同使用人）
    app_accounts_set = set(app_admin["账号名"].dropna().astype(str).str.strip())
    db_accounts_set = set(db_admin["账号名"].dropna().astype(str).str.strip())
    common_accounts = app_accounts_set & db_accounts_set

    skip_accounts = {"", "root", "ops", "dev", "pam_root", "halt", "sync", "shutdown"}
    existing_users = {r["重合人"] for r in rows}

    for acct in sorted(common_accounts):
        if acct in skip_accounts:
            continue

        app_records = app_admin[app_admin["账号名"].astype(str).str.strip() == acct]
        db_records = db_admin[db_admin["账号名"].astype(str).str.strip() == acct]

        app_users_for_acct = set(app_records["账号使用人"].astype(str).str.strip())
        db_users_for_acct = set(db_records["账号使用人"].astype(str).str.strip())
        all_users = app_users_for_acct | db_users_for_acct

        # 跳过已由「使用人」匹配到的
        if all_users & existing_users:
            continue

        app_objects = set(app_records["审阅对象"].astype(str).str.strip())
        db_objects = set(db_records["审阅对象"].astype(str).str.strip())

        detail = (
            f"应用层: {'/'.join(sorted(app_objects))} "
            f"(使用人: {'/'.join(sorted(app_users_for_acct))}); "
            f"数据库/服务器: {'/'.join(sorted(db_objects))} "
            f"(使用人: {'/'.join(sorted(db_users_for_acct))})"
        )

        rows.append({
            "系统": system,
            "重合人": " / ".join(sorted(all_users)),
            "重合账号": acct,
            "具体哪个重合": detail,
        })

    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    all_permission_results: List[pd.DataFrame] = []
    all_overlap_results: List[pd.DataFrame] = []

    for system in SYSTEMS:
        print(f"\n{'='*60}")
        print(f"处理系统: {system}")
        print(f"{'='*60}")

        # 1. 读取数据
        dev_df = read_developer_list(system)
        audit_df = read_audit_export(system)

        print(f"  开发人员: {len(dev_df)} 人")
        print(f"  审阅记录: {len(audit_df)} 条")

        # 2. 权限检查
        perm_df = check_dev_permissions(system, dev_df, audit_df)
        all_permission_results.append(perm_df)

        readonly_ok = len(perm_df[perm_df["权限是否只读"] == "是"])
        readonly_ng = len(perm_df[perm_df["权限是否只读"] == "否"])
        no_match = len(perm_df[perm_df["权限是否只读"] == "未找到匹配记录"])
        print(f"  权限检查: 只读={readonly_ok}, 非只读={readonly_ng}, 未匹配={no_match}")

        # 3. 管理员重合检查
        overlap_df = find_admin_overlap(system, audit_df)
        all_overlap_results.append(overlap_df)

        print(f"  管理员重合: {len(overlap_df)} 条")
        for _, row in overlap_df.iterrows():
            print(f"    → {row['重合人']} | {row['重合账号']}")

    # 4. 输出合并结果
    final_perms = pd.concat(all_permission_results, ignore_index=True)
    final_overlaps = pd.concat(all_overlap_results, ignore_index=True)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        if not final_perms.empty:
            final_perms = final_perms.fillna("")
            final_perms.to_excel(writer, sheet_name="权限检查", index=False)
        else:
            pd.DataFrame({"提示": ["无数据"]}).to_excel(writer, sheet_name="权限检查", index=False)

        if not final_overlaps.empty:
            final_overlaps = final_overlaps.fillna("")
            final_overlaps.to_excel(writer, sheet_name="管理员重合", index=False)
        else:
            pd.DataFrame({"提示": ["未发现重合"]}).to_excel(writer, sheet_name="管理员重合", index=False)

    print(f"\n✅ 输出文件: {OUTPUT_FILE}")
    print(f"   - Sheet「权限检查」: {len(final_perms)} 行")
    print(f"   - Sheet「管理员重合」: {len(final_overlaps)} 行")


if __name__ == "__main__":
    main()
