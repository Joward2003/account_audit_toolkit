from pathlib import Path
from typing import Dict, List, Optional, Tuple
import re
import difflib

import pandas as pd


# ============================================================
# 1. 基础工具函数
# ============================================================

def normalize_key(value) -> str:
    """
    用于匹配 system / folder 的标准化函数。

    作用：
    - 去空
    - 转大写
    - 去掉空格、下划线、中划线、斜杠等干扰符号

    例如：
    PMS/VMS、PMS_VMS、pms vms
    都会变成 PMSVMS
    """
    if value is None or pd.isna(value):
        return ""

    text = str(value).strip().upper()
    text = re.sub(r"[\s_\-/\\（）()&]+", "", text)
    return text


def scan_folders(supporting_root: str | Path) -> pd.DataFrame:
    """
    扫描 supporting 根目录下的一级文件夹。

    例如：
    supporting_root/
      ├── BMS/
      ├── VMall/
      ├── PMS_VMS公共凭证/
      └── MaYun/
    """
    root = Path(supporting_root)

    if not root.exists():
        raise FileNotFoundError(f"supporting 根目录不存在：{root}")

    rows = []

    for p in root.iterdir():
        if p.is_dir():
            rows.append(
                {
                    "folder_name": p.name,
                    "folder_path": str(p.resolve()),
                    "folder_key": normalize_key(p.name),
                }
            )

    return pd.DataFrame(rows)


def load_mapping(mapping_path: str | Path) -> pd.DataFrame:
    """
    读取 system_folder_mapping.xlsx。

    需要至少包含：
    - system
    - actual_folder_name

    可选包含：
    - alias
    - note
    """
    mapping_path = Path(mapping_path)

    if not mapping_path.exists():
        return pd.DataFrame(
            columns=["system", "actual_folder_name", "alias", "note"]
        )

    df = pd.read_excel(mapping_path)

    required_cols = {"system", "actual_folder_name"}
    missing = required_cols - set(df.columns)

    if missing:
        raise ValueError(
            f"{mapping_path} 缺少必要列：{missing}。"
            "请至少保留 system 和 actual_folder_name 两列。"
        )

    if "alias" not in df.columns:
        df["alias"] = ""

    if "note" not in df.columns:
        df["note"] = ""

    return df


# ============================================================
# 2. 构建映射字典
# ============================================================

def build_manual_mapping(mapping_df: pd.DataFrame) -> Dict[str, Dict[str, str]]:
    """
    把 system_folder_mapping.xlsx 转成字典。

    返回结果类似：
    {
        "PMSVMS": {
            "system": "PMS/VMS",
            "actual_folder_name": "PMS_VMS公共凭证",
            "note": "公共平台"
        }
    }
    """
    mapping = {}

    for _, row in mapping_df.iterrows():
        system = row.get("system", "")
        actual_folder_name = row.get("actual_folder_name", "")
        note = row.get("note", "")

        if pd.isna(system) or str(system).strip() == "":
            continue

        key = normalize_key(system)

        mapping[key] = {
            "system": str(system).strip(),
            "actual_folder_name": "" if pd.isna(actual_folder_name) else str(actual_folder_name).strip(),
            "note": "" if pd.isna(note) else str(note).strip(),
        }

        # alias 支持用英文分号 ; 或中文分号 ； 分隔
        alias_text = row.get("alias", "")
        if not pd.isna(alias_text) and str(alias_text).strip():
            aliases = re.split(r"[;；]", str(alias_text))

            for alias in aliases:
                alias_key = normalize_key(alias)
                if alias_key:
                    mapping[alias_key] = {
                        "system": str(system).strip(),
                        "actual_folder_name": "" if pd.isna(actual_folder_name) else str(actual_folder_name).strip(),
                        "note": f"alias of {system}",
                    }

    return mapping


def build_folder_index(folder_df: pd.DataFrame) -> Dict[str, Dict[str, str]]:
    """
    把实际文件夹列表转成字典。

    返回结果类似：
    {
        "BMS": {
            "folder_name": "BMS",
            "folder_path": "/xxx/supporting_root/BMS"
        }
    }
    """
    folder_index = {}

    for _, row in folder_df.iterrows():
        folder_key = row["folder_key"]

        folder_index[folder_key] = {
            "folder_name": row["folder_name"],
            "folder_path": row["folder_path"],
        }

    return folder_index


# ============================================================
# 3. 单个 TargetSystem 匹配文件夹
# ============================================================

def resolve_folder_for_system(
    target_system: str,
    manual_mapping: Dict[str, Dict[str, str]],
    folder_index: Dict[str, Dict[str, str]],
) -> Dict[str, str]:
    """
    根据 TargetSystem 找 supporting 文件夹。

    优先级：
    1. 先查 system_folder_mapping.xlsx 人工映射
    2. 如果没有人工映射，则默认 TargetSystem = 文件夹名
    3. 如果还是找不到，则给一个 fuzzy suggestion，但不自动通过
    """
    target_key = normalize_key(target_system)

    # 1. 先查人工映射表
    if target_key in manual_mapping:
        mapped = manual_mapping[target_key]
        actual_folder_name = mapped["actual_folder_name"]
        actual_folder_key = normalize_key(actual_folder_name)

        if actual_folder_key in folder_index:
            folder = folder_index[actual_folder_key]
            return {
                "ResolvedFolderName": folder["folder_name"],
                "ResolvedFolderPath": folder["folder_path"],
                "FolderMatchStatus": "matched_manual_mapping",
                "FolderMatchNote": mapped.get("note", ""),
            }

        return {
            "ResolvedFolderName": actual_folder_name,
            "ResolvedFolderPath": "",
            "FolderMatchStatus": "mapping_exists_but_folder_missing",
            "FolderMatchNote": f"mapping 表里写了 {actual_folder_name}，但实际目录下没找到这个文件夹",
        }

    # 2. 默认 TargetSystem = 文件夹名
    if target_key in folder_index:
        folder = folder_index[target_key]
        return {
            "ResolvedFolderName": folder["folder_name"],
            "ResolvedFolderPath": folder["folder_path"],
            "FolderMatchStatus": "matched_default",
            "FolderMatchNote": "默认匹配：TargetSystem 与文件夹名一致",
        }

    # 3. 找不到时，只给建议，不自动通过
    folder_keys = list(folder_index.keys())
    suggestions = difflib.get_close_matches(target_key, folder_keys, n=3, cutoff=0.65)

    if suggestions:
        suggested_names = [
            folder_index[k]["folder_name"]
            for k in suggestions
        ]

        return {
            "ResolvedFolderName": "",
            "ResolvedFolderPath": "",
            "FolderMatchStatus": "review_needed_with_suggestion",
            "FolderMatchNote": "可能匹配：" + " | ".join(suggested_names),
        }

    return {
        "ResolvedFolderName": "",
        "ResolvedFolderPath": "",
        "FolderMatchStatus": "unmatched",
        "FolderMatchNote": "未找到对应 supporting 文件夹，需要维护 system_folder_mapping.xlsx",
    }


# ============================================================
# 4. 批量处理 todo_list
# ============================================================

def resolve_todo_folders(
    todo_path: str | Path,
    supporting_root: str | Path,
    mapping_path: str | Path = "configs/system_folder_mapping.xlsx",
    output_path: str | Path = "output/todo_list_with_folder.xlsx",
) -> Dict[str, str]:
    """
    主入口函数。

    输入：
    - todo_path：Scope Engine 生成的 Excel，里面要有 todo_list sheet
    - supporting_root：所有系统 supporting 文件夹所在的根目录
    - mapping_path：system_folder_mapping.xlsx
    - output_path：输出结果

    输出：
    - todo_list_with_folder.xlsx
    - unmatched_systems.xlsx
    """
    todo_path = Path(todo_path)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # 读取 todo_list
    todo_df = pd.read_excel(todo_path, sheet_name="todo_list")

    if "TargetSystem" not in todo_df.columns:
        if "System" not in todo_df.columns:
            raise ValueError("todo_list 中缺少 TargetSystem 和 System，无法定位系统。")
        todo_df["TargetSystem"] = todo_df["System"]

    # 扫描实际文件夹
    folder_df = scan_folders(supporting_root)

    # 读取人工映射表
    mapping_df = load_mapping(mapping_path)

    # 构建字典
    manual_mapping = build_manual_mapping(mapping_df)
    folder_index = build_folder_index(folder_df)

    # 逐行解析
    resolved_rows: List[Dict[str, str]] = []

    for _, row in todo_df.iterrows():
        target_system = row["TargetSystem"]

        resolved = resolve_folder_for_system(
            target_system=target_system,
            manual_mapping=manual_mapping,
            folder_index=folder_index,
        )

        resolved_rows.append(resolved)

    resolved_df = pd.DataFrame(resolved_rows)

    final_df = pd.concat(
        [todo_df.reset_index(drop=True), resolved_df],
        axis=1,
    )

    # 生成未匹配系统清单
    unmatched_df = (
        final_df[
            final_df["FolderMatchStatus"].isin(
                [
                    "unmatched",
                    "review_needed_with_suggestion",
                    "mapping_exists_but_folder_missing",
                ]
            )
        ][
            [
                "TargetSystem",
                "ResolvedFolderName",
                "FolderMatchStatus",
                "FolderMatchNote",
            ]
        ]
        .drop_duplicates()
        .reset_index(drop=True)
    )

    unmatched_path = output_path.parent / "unmatched_systems.xlsx"

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        final_df.to_excel(writer, sheet_name="todo_with_folder", index=False)
        folder_df.to_excel(writer, sheet_name="scanned_folders", index=False)
        mapping_df.to_excel(writer, sheet_name="mapping_used", index=False)
        unmatched_df.to_excel(writer, sheet_name="unmatched_systems", index=False)

    unmatched_df.to_excel(unmatched_path, index=False)

    return {
        "todo_with_folder": str(output_path),
        "unmatched_systems": str(unmatched_path),
    }


# ============================================================
# 5. 命令行运行
# ============================================================

if __name__ == "__main__":
    """
    示例运行方式：

    python system_folder_mapper.py \
        --todo output/account_audit_todo_demo.xlsx \
        --supporting-root data/supporting_root \
        --mapping configs/system_folder_mapping.xlsx \
        --output output/todo_list_with_folder.xlsx
    """

    import argparse
    import json

    parser = argparse.ArgumentParser(
        description="Map TargetSystem in todo_list to actual supporting folders."
    )

    parser.add_argument(
        "--todo",
        required=True,
        help="Scope Engine 生成的 Excel 文件路径，要求包含 todo_list sheet。",
    )

    parser.add_argument(
        "--supporting-root",
        required=True,
        help="所有系统 supporting 文件夹所在的根目录。",
    )

    parser.add_argument(
        "--mapping",
        default="configs/system_folder_mapping.xlsx",
        help="system_folder_mapping.xlsx 路径。",
    )

    parser.add_argument(
        "--output",
        default="output/todo_list_with_folder.xlsx",
        help="输出文件路径。",
    )

    args = parser.parse_args()

    files = resolve_todo_folders(
        todo_path=args.todo,
        supporting_root=args.supporting_root,
        mapping_path=args.mapping,
        output_path=args.output,
    )

    print(json.dumps(files, ensure_ascii=False, indent=2))