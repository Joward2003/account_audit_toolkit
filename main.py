"""Run account audit workflow export — 完整版。

端到端流水线，将 scope engine、execution plan builder、evidence checker 串联，
生成一份面向业务阅读的账号审阅结论 Excel。

流水线:
  WP Excel (test.xlsx)
    → Scope Engine (哪些节点要查)
    → System-Folder Mapper (映射到 supporting 文件夹)
    → Execution Plan Builder (build_expected_accounts + decide_evidence_plan)
    → Evidence Checker (截图文字比对，仅 支持性材料标识=Y 时)
    → 最终审阅结论 Excel（范围、结果、分类明细）

输出:
  output/account_audit_conclusion.xlsx
    - Sheet「测试范围」: 每个系统测不测、怎么测、判断依据
    - Sheet「测试结果」: 每个系统每项测试的结果和说明
    - Sheet「明细-截图核对」: 仅展示真的执行截图核对的账号逐项结果
    - Sheet「明细-导出核对」: 展示只查账号审阅导出的账号清单
    - Sheet「明细-开发人员」: 展示开发人员清单和权限核对
    - Sheet「明细-SOD」: 展示前后台管理员重合核对

用法:
  python main.py
  python main.py --input ../test.xlsx --sheet Sheet1
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

import pandas as pd

from audit_nodes import AuditNodes
from config import AuditConfig
from exporter import AccountAuditWorkflowExporter
from pipeline import AccountAuditPipeline

# ---- 直接复用 execution_plan_builder 的完整逻辑 ----
from execution_plan_builder import (
    find_account_export_file,
    read_account_export,
    build_expected_accounts_for_task,
    decide_evidence_plan,
    NODE_TO_REVIEW_OBJECT,
    EVIDENCE_ONLY_NODE_REVIEW_OBJECT,
    AccountExportReadResult,
    EXPORT_ACCOUNT_CANDIDATES,
    EXPORT_OWNER_CANDIDATES,
    EXPORT_USAGE_CANDIDATES,
    resolve_column,
    split_multi_values,
)

# ---- 直接复用 evidence_checker ----
from evidence_checker import (
    EvidenceLoader,
    OCRExtractor,
    EvidenceChecker,
    FULL_NODE_TO_SHEET,
)
from system_folder_mapper import build_manual_mapping, load_mapping, normalize_key
from sid_utils import normalize_sid
from application_admin_role_config import (
    load_application_admin_role_rules,
    match_role_keyword,
)


# ============================================================
# 配置
# ============================================================

APP_DIR = Path(__file__).resolve().parent
PROJECT_DIR = APP_DIR.parent
SUPPORTING_DIR = (
    APP_DIR / "supporting"
    if (APP_DIR / "supporting").exists()
    else PROJECT_DIR / "supporting"
)
OUTPUT_DIR = APP_DIR / "output"
DEFAULT_INPUT = (
    APP_DIR / "test.xlsx"
    if (APP_DIR / "test.xlsx").exists()
    else PROJECT_DIR / "test.xlsx"
)
CONCLUSION_FILE = OUTPUT_DIR / "account_audit_conclusion.xlsx"
CONFIG_DIR = APP_DIR / "configs"
SYSTEM_FOLDER_MAPPING_FILE = CONFIG_DIR / "系统文件夹映射表.xlsx"
LEGACY_SYSTEM_FOLDER_MAPPING_FILE = CONFIG_DIR / "system_folder_mapping.xlsx"
NAME_ALIAS_FILE = CONFIG_DIR / "中英文姓名对照表.xlsx"
LEGACY_NAME_ALIAS_FILE = CONFIG_DIR / "name_alias_mapping.xlsx"
ACCOUNT_ALIAS_FILE = CONFIG_DIR / "账号别名映射表.xlsx"
LEGACY_ACCOUNT_ALIAS_FILE = CONFIG_DIR / "account_alias_mapping.xlsx"


# ============================================================
# 主类
# ============================================================

class AuditConclusionGenerator:
    """端到端审阅结论生成器。"""

    def __init__(self):
        self.config = AuditConfig()
        self.pipeline = AccountAuditPipeline(self.config)
        self.exporter = AccountAuditWorkflowExporter(self.config)
        self.evidence_loader = EvidenceLoader()
        self.ocr_extractor = OCRExtractor()
        self.evidence_checker = EvidenceChecker()

    # ----------------------------------------------------------
    # 主入口
    # ----------------------------------------------------------

    def generate(
        self,
        wp_df: pd.DataFrame,
        supporting_dir: Optional[Path] = None,
        progress_callback: Optional[Callable[[Dict[str, Any]], None]] = None,
    ) -> Dict[str, Any]:
        def emit(stage: str, message: str, **detail: Any) -> None:
            if progress_callback:
                progress_callback({"stage": stage, "message": message, **detail})

        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        _supporting = supporting_dir or SUPPORTING_DIR

        # Step 1: Scope Engine
        emit("scope", "正在分析各系统适用的审阅测试")
        print("=" * 60)
        print("Step 1: Scope Engine")
        print("=" * 60)
        scope_results = self.exporter.build_scope_results(wp_df)
        self._validate_scope_sid_identity(scope_results)
        todo_df = self.exporter.build_todo_df(scope_results)
        detail_df = self.pipeline.run(wp_df)
        for _, row in wp_df.iterrows():
            print(f"  {row['System']:15s} | {row['Category']}")

        # Step 2: Folder mapping
        emit("folder_mapping", "正在识别支持性文件目录并匹配系统")
        print("\n" + "=" * 60)
        print("Step 2: Folder Mapper")
        print("=" * 60)
        system_entries = [
            (normalize_sid(result.facts.sid), str(result.facts.system).strip())
            for result in scope_results
        ]
        folder_map = self._build_folder_map(
            _supporting, [sid for sid, _ in system_entries]
        )
        if supporting_dir:
            print(f"  [使用上传的 supporting: {_supporting}]", flush=True)
            print(f"  [文件夹映射: { {k: Path(v).name for k, v in folder_map.items()} }]", flush=True)
        for sid, sys_name in system_entries:
            fp = self._resolve_folder(sid, folder_map)
            print(f"  {str(sys_name):15s} [SID={sid or '缺失'}] → {fp or 'NOT FOUND'}")

        # Step 3: Per-system check
        print("\n" + "=" * 60)
        print("Step 3: Evidence Check（账号审阅导出 + 截图OCR仅Y时）")
        print("=" * 60)

        scope_baseline = self._build_scope_baseline(scope_results)
        system_results: Dict[str, Dict[str, Dict]] = {}
        for sid, system_str in system_entries:
            emit(
                "system",
                f"正在检查系统：{system_str}",
                system=system_str,
                sid=sid,
            )
            folder_path = self._resolve_folder(sid, folder_map)
            system_results[sid] = self._check_system(
                system_str,
                sid,
                folder_path,
                todo_df,
                scope_baseline.get(sid, {}),
                progress_callback,
            )

        # Step 4: Build conclusion Excel
        print("\n" + "=" * 60)
        print("Step 4: 生成审阅结论 Excel")
        print("=" * 60)
        emit("export", "正在汇总检查结果并生成结论 Excel")
        test_scope_df = self._build_test_scope_df(wp_df, system_results)
        test_result_df = self._build_test_result_df(wp_df, system_results)
        detail_sheets = self._build_business_detail_sheets(wp_df, system_results)

        with pd.ExcelWriter(CONCLUSION_FILE, engine="openpyxl") as writer:
            test_scope_df.to_excel(writer, sheet_name="测试范围", index=False)
            test_result_df.to_excel(writer, sheet_name="测试结果", index=False)
            for sheet_name, detail_df in detail_sheets.items():
                detail_df.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"  ✅ {CONCLUSION_FILE}")
        emit("completed", "检查完成")
        return {"conclusion": str(CONCLUSION_FILE)}

    # ----------------------------------------------------------
    # Folder mapping
    # ----------------------------------------------------------

    @staticmethod
    def _build_folder_map(root_dir: Path, sids: List[str]) -> Dict[str, str]:
        """递归建立 SID -> supporting 文件夹索引，并支持单系统 ZIP 直接放 Excel。"""
        m: Dict[str, str] = {}
        if not root_dir.exists():
            return m

        directories = [root_dir]
        directories.extend(
            p for p in root_dir.rglob("*")
            if p.is_dir()
            and "__MACOSX" not in p.parts
            and not p.name.startswith(".")
            and not p.name.startswith("._")
        )
        for d in directories:
            if d == root_dir:
                continue
            m.setdefault(normalize_key(d.name), str(d))

        # 接入 SID 配置表，不再在代码中维护系统名称别名。
        mapping_path = SYSTEM_FOLDER_MAPPING_FILE if SYSTEM_FOLDER_MAPPING_FILE.exists() else LEGACY_SYSTEM_FOLDER_MAPPING_FILE
        manual_mapping = build_manual_mapping(load_mapping(mapping_path))
        by_folder_name = {normalize_key(Path(path).name): path for path in m.values()}
        for sid_key, rule in manual_mapping.items():
            actual_key = normalize_key(rule.get("actual_folder_name", ""))
            if actual_key in by_folder_name:
                m[sid_key] = by_folder_name[actual_key]

        # 单系统压缩包允许直接把 Excel 放在根目录或任意一层包装目录。
        unique_sids = sorted({sid for sid in sids if sid})
        if len(unique_sids) == 1:
            sid_key = normalize_sid(unique_sids[0])
            if sid_key not in m:
                content_dirs = [
                    d for d in directories
                    if any(
                        p.is_file() and p.suffix.lower() in {".xlsx", ".xls", ".csv"}
                        and not p.name.startswith("~$")
                        for p in d.iterdir()
                    )
                ]
                if content_dirs:
                    content_dirs.sort(key=lambda p: (len(p.relative_to(root_dir).parts), str(p)))
                    m[sid_key] = str(content_dirs[0])
        return m

    @staticmethod
    def _resolve_folder(sid: str, folder_map: Dict[str, str]) -> str:
        key = normalize_sid(sid)
        if not key:
            return ""
        return folder_map.get(key, "")

    @staticmethod
    def _validate_scope_sid_identity(scope_results: List[Any]) -> None:
        """SID 是运行时唯一键；缺失或重复都会导致结果串系统，因此直接阻断。"""
        seen: Dict[str, str] = {}
        for result in scope_results:
            sid = normalize_sid(result.facts.sid)
            system = str(result.facts.system).strip()
            if not sid:
                raise ValueError(f"系统“{system or '未命名'}”缺少 SID，无法执行 SID 映射。")
            if sid in seen:
                raise ValueError(
                    f"WP 中 SID 重复：{sid} 同时出现在“{seen[sid]}”和“{system}”。"
                    "请确保每一行系统使用唯一 SID。"
                )
            seen[sid] = system

    @staticmethod
    def _build_scope_baseline(scope_results: List[Any]) -> Dict[str, Dict[str, Dict[str, Any]]]:
        """保留所有节点的 Scope 结论，避免 N/A/待确认节点在最终表中变成“未评估”。"""
        baseline: Dict[str, Dict[str, Dict[str, Any]]] = {}
        for result in scope_results:
            sid = normalize_sid(result.facts.sid)
            baseline[sid] = {}
            warning = " | ".join(result.warnings)
            for node in AuditNodes.ALL:
                decision = result.decisions.get(node)
                if decision is None:
                    continue
                if decision.label.startswith("N/A"):
                    final_status = "不适用"
                elif decision.label.startswith("Refer to"):
                    final_status = "引用"
                elif decision.label == "待确认":
                    final_status = "待确认"
                else:
                    final_status = "待执行"
                baseline[sid][node] = {
                    "scope": decision.label,
                    "scope_reason": decision.reason,
                    "scope_warning": warning,
                    "final": decision.label,
                    "final_reason": decision.reason,
                    "final_status": final_status,
                }
        return baseline

    @staticmethod
    def _missing_folder_result(task: pd.Series) -> Dict[str, Any]:
        reason = (
            f"未找到系统“{task.get('TargetSystem', task.get('System', ''))}”"
            "对应的支持性文件夹。"
        )
        return {
            "scope": str(task.get("Scope结果", "")),
            "scope_reason": str(task.get("Scope原因", "")),
            "scope_warning": str(task.get("Scope警告", "")),
            "account_ok": "未检查",
            "manual_upload": "UNKNOWN",
            "ocr": "未执行",
            "ocr_needed": "否",
            "ocr_trigger_reason": "支持性文件夹未匹配，无法判断是否需要核对截图",
            "final": f"否-{reason}",
            "final_reason": reason,
            "final_status": "未通过",
        }

    @staticmethod
    def _ocr_not_triggered_reason(parse_method: str, manual_overall: str) -> str:
        if manual_overall == "N" and parse_method == "no_manual_upload":
            return "手工上传标识=N，该节点不要求上传截图支持文件"
        if parse_method == "export_available":
            return "该节点使用系统导出检查，不要求截图"
        if manual_overall == "PARTIAL":
            return "手工上传标识同时存在Y和N，需先人工确认"
        if manual_overall in {"UNKNOWN", "N_OR_REVIEW", ""}:
            return "支持性材料标识缺失或无法识别，暂未核对截图"
        return f"检查方式={parse_method or '未配置'}，暂未核对截图"

    # ----------------------------------------------------------
    # 文件查找（兼容 zip 编码损坏）
    # ----------------------------------------------------------

    @staticmethod
    def _find_file_by_content(folder: Path, filename_hint: str,
                              required_columns: List[str]) -> Optional[Path]:
        """
        在文件夹中查找文件。先按文件名匹配，失败后按内容匹配。
        用于兼容 zip 包中中文文件名编码损坏的场景。
        """
        # 先按文件名精确匹配
        exact = folder / filename_hint
        if exact.exists():
            return exact
        # rglob 搜索
        for p in folder.rglob(filename_hint):
            if p.is_file():
                return p

        # 按内容匹配：扫描所有 xlsx，找包含 required_columns 的
        for p in sorted(folder.rglob("*.xlsx"), key=lambda x: len(x.parts)):
            if p.name.startswith("~$"):
                continue
            try:
                df_sample = pd.read_excel(p, nrows=0)
                cols = set(str(c) for c in df_sample.columns)
                if cols.issuperset(set(required_columns)):
                    return p
                # 也检查所有 sheet
                xl = pd.ExcelFile(p)
                for s in xl.sheet_names[:3]:
                    df_s = pd.read_excel(p, sheet_name=s, nrows=0)
                    cols_s = set(str(c) for c in df_s.columns)
                    if cols_s.issuperset(set(required_columns)):
                        return p
            except Exception:
                continue
        return None

    # ----------------------------------------------------------
    # 单系统检查
    # ----------------------------------------------------------

    def _check_system(
        self,
        system: str,
        sid: str,
        folder_path: str,
        todo_df: pd.DataFrame,
        baseline: Dict[str, Dict[str, Any]],
        progress_callback: Optional[Callable[[Dict[str, Any]], None]] = None,
    ) -> Dict[str, Dict]:
        print(f"\n--- {system} ---")
        print(f"    folder_path={folder_path}", flush=True)

        conclusions: Dict[str, Dict] = {
            node: dict(result) for node, result in baseline.items()
        }

        # 获取该系统任务
        system_todo = (
            todo_df[todo_df["SID"].map(normalize_sid) == normalize_sid(sid)]
            if "SID" in todo_df.columns
            else pd.DataFrame()
        )

        if not folder_path:
            print("  ⚠ 无文件夹映射")
            if progress_callback:
                progress_callback({
                    "stage": "folder_missing",
                    "message": f"未找到系统支持性文件夹：{system}",
                    "system": system,
                })
            for _, task in system_todo.iterrows():
                node = str(task.get("检查节点", ""))
                conclusions[node] = self._missing_folder_result(task)
            return conclusions

        # 读取账号审阅导出
        export_file = find_account_export_file(folder_path)
        if export_file is None and folder_path:
            _contents = list(Path(folder_path).rglob("*.xlsx"))
            print(f"    folder 中所有 xlsx: {[str(p.relative_to(folder_path)) for p in _contents]}", flush=True)
        export_result = (
            read_account_export(export_file) if export_file
            else AccountExportReadResult(status="ACCOUNT_EXPORT_FILE_NOT_FOUND")
        )
        print(f"  账号审阅导出: {export_file.name if export_file else 'NOT FOUND'} "
              f"[{export_result.status}]")

        # 找 supporting 截图 Excel
        supporting_excel = Path(folder_path) / "支持性附件截图.xlsx"
        if not supporting_excel.exists():
            # 文件名损坏兜底：找既不是账号导出也不是开发人员清单的 xlsx
            for p in sorted(Path(folder_path).rglob("*.xlsx"), key=lambda x: len(x.parts)):
                if p.name.startswith("~$"):
                    continue
                try:
                    df_t = pd.read_excel(p, nrows=0)
                    cols = set(str(c) for c in df_t.columns)
                    if "审阅对象" in cols or "账号名" in cols:
                        continue
                    if "姓名" in cols or "域账号" in cols:
                        continue
                except Exception:
                    pass
                supporting_excel = p
                break
        if not supporting_excel.exists():
            supporting_excel = None

        for _, task in system_todo.iterrows():
            node = str(task.get("检查节点", ""))
            scope_label = str(task.get("Scope结果", ""))
            if progress_callback:
                progress_callback({
                    "stage": "node",
                    "message": f"正在检查：{system} / {node}",
                    "system": system,
                    "node": node,
                })
            conclusions[node] = self._check_one_node(
                node,
                scope_label,
                task,
                export_result,
                folder_path,
                supporting_excel,
                sid,
                progress_callback,
            )

        return conclusions

    def _check_one_node(
        self,
        node: str,
        scope_label: str,
        task: pd.Series,
        export_result: AccountExportReadResult,
        folder_path: str,
        supporting_excel: Optional[Path],
        sid: str,
        progress_callback: Optional[Callable[[Dict[str, Any]], None]] = None,
    ) -> Dict[str, Any]:
        """对单个审计节点执行完整测试逻辑。"""

        # Scope 判定：N/A / Refer to 直接返回，不做本系统账号检查。
        # 其他进入 todo 的自定义结论同样必须执行，不能因标签不是“是”而跳过。
        scope_label_str = str(scope_label)
        if scope_label_str.startswith("Refer to"):
            result = {
                "scope": scope_label_str,
                "scope_reason": str(task.get("Scope原因", "")),
                "account_ok": "N/A",
                "manual_upload": "N/A",
                "ocr": "N/A",
                "final": scope_label_str,
                "final_reason": str(task.get("Scope原因", "")),
                "final_status": "引用",
            }
            print(f"  [{node}] scope={scope_label_str} → {result['final']}")
            return result

        if scope_label_str.startswith("N/A"):
            result = {
                "scope": scope_label_str,
                "scope_reason": str(task.get("Scope原因", "")),
                "account_ok": "N/A",
                "manual_upload": "N/A",
                "ocr": "N/A",
                "final": scope_label_str if scope_label_str else "N/A",
                "final_reason": str(task.get("Scope原因", "")),
                "final_status": "不适用",
            }
            print(f"  [{node}] scope={scope_label_str} → {result['final']}")
            return result

        # ---- 1. 调用 execution_plan_builder 构建 expected accounts ----
        review_object = NODE_TO_REVIEW_OBJECT.get(node, "")
        if node in EVIDENCE_ONLY_NODE_REVIEW_OBJECT:
            review_object = EVIDENCE_ONLY_NODE_REVIEW_OBJECT[node]
        _, summary = build_expected_accounts_for_task(export_result, task, review_object)

        expected_count = summary.get("ExpectedCount", 0)
        expected_status = summary.get("ExpectedSetStatus", "")
        expected_note = str(summary.get("ExpectedSetNote", ""))
        manual_overall = str(summary.get("ManualUploadOverall", "UNKNOWN")).strip()

        # ---- 2. 调用 execution_plan_builder 判断证据计划 ----
        plan = decide_evidence_plan(task, summary)
        parse_method = plan.get("ParseMethod", "")
        evidence_plan_status = str(plan.get("EvidencePlanStatus", ""))
        evidence_plan_note = str(plan.get("EvidencePlanNote", ""))

        # ---- 3. 账号存在性检查 ----
        if expected_count > 0:
            account_ok = True
            account_msg = f"通过({expected_count})"
        elif expected_status in {"ONLY_SKIPPED_ACCOUNTS"}:
            account_ok = False
            account_msg = "仅系统账号"
        elif expected_status == "NO_REVIEW_OBJECT_RULE":
            account_ok = None
            account_msg = "N/A"
        else:
            account_ok = False
            account_msg = f"未通过({expected_status})"

        # ---- 4. 根据 evidence plan 决定是否需要 OCR ----
        # 只有 parse_method 指向需要查截图/OCR 的才跑
        OCR_PARSE_METHODS = {
            "account_presence_or_matrix_review",
            "server_evidence_review_from_manual_upload",
            "evidence_exists_from_manual_upload_flag",
        }
        ocr_needed = parse_method in OCR_PARSE_METHODS and manual_overall == "Y"
        ocr_ok: Optional[bool] = None
        ocr_msg = "无需核对截图"
        ocr_detail = ""
        ocr_recognized = ""
        ocr_matched = ""
        ocr_global_matched = ""
        ocr_missing = ""
        ocr_image_count: Any = ""
        ocr_trigger_reason = (
            "手工上传标识=Y且该检查方式需要截图核验"
            if ocr_needed
            else self._ocr_not_triggered_reason(parse_method, manual_overall)
        )

        if ocr_needed:
            if progress_callback:
                progress_callback({
                    "stage": "ocr",
                    "message": f"正在核对截图文字：{task.get('System', '')} / {node}",
                    "system": str(task.get("System", "")),
                    "node": node,
                })
            ocr_result = self._run_ocr_check(node, task, supporting_excel, summary)
            ocr_ok = ocr_result.get("status", "") == "通过"
            ocr_msg = ocr_result.get("status", "截图核对失败")
            ocr_detail = ocr_result.get("detail", "")
            ocr_recognized = ocr_result.get("ocr_recognized_accounts", "")
            ocr_matched = ocr_result.get("matched_accounts", "")
            ocr_global_matched = ocr_result.get("global_matched_accounts", "")
            ocr_missing = ocr_result.get("missing_accounts", "")
            ocr_image_count = ocr_result.get("image_count", "")

        # ---- 5. 派生检查（开发人员清单/SOD） ----
        derived_result = self._check_derived_nodes(
            node, folder_path, export_result, supporting_excel, sid
        )

        # ---- 6. 汇总 ----
        final, final_reason, final_status = self._compute_final(
            account_ok=account_ok,
            account_detail=expected_note or account_msg,
            manual_overall=manual_overall,
            parse_method=parse_method,
            ocr_needed=ocr_needed,
            ocr_ok=ocr_ok,
            ocr_status=ocr_msg,
            ocr_detail=ocr_detail,
            derived_result=derived_result,
        )

        print(f"  [{node}] account={account_msg} | upload={manual_overall} | "
              f"parse={parse_method} | ocr={'Y' if ocr_needed else 'N'}={ocr_msg}"
              f"{' | derived=' + derived_result.get('status','?') if derived_result else ''}"
              f" | → {final}")

        return {
            "scope": scope_label,
            "scope_reason": str(task.get("Scope原因", "")),
            "scope_warning": str(task.get("Scope警告", "")),
            "account_ok": account_msg,
            "expected_status": expected_status,
            "expected_note": expected_note,
            "expected_count": expected_count,
            "expected_accounts": str(summary.get("ExpectedAccounts", "")),
            "review_object": review_object,
            "sid": sid,
            "export_file": str(summary.get("ExportFile", "")),
            "export_sheet": str(summary.get("ExportSheet", "")),
            "manual_upload": manual_overall,
            "manual_upload_note": str(summary.get("ManualUploadNote", "")),
            "ocr": ocr_msg,
            "ocr_needed": "是" if ocr_needed else "否",
            "ocr_trigger_reason": ocr_trigger_reason,
            "ocr_detail": ocr_detail,
            "ocr_recognized_accounts": ocr_recognized,
            "ocr_image_count": ocr_image_count,
            "ocr_matched_accounts": ocr_matched,
            "ocr_global_matched_accounts": ocr_global_matched,
            "ocr_missing_accounts": ocr_missing,
            "supporting_file": str(supporting_excel or ""),
            "supporting_sheet": str(plan.get("SupportingSheet", "")),
            "parse_method": parse_method,
            "evidence_plan_status": evidence_plan_status,
            "evidence_plan_note": evidence_plan_note,
            "derived_status": derived_result.get("status", "") if derived_result else "",
            "derived_detail": derived_result.get("detail", "") if derived_result else "",
            "derived_records": derived_result.get("records", []) if derived_result else [],
            "final": final,
            "final_reason": final_reason,
            "final_status": final_status,
        }

    # ----------------------------------------------------------
    # OCR 检查（仅在需要时调用）
    # ----------------------------------------------------------

    def _run_ocr_check(
        self,
        node: str,
        task: pd.Series,
        supporting_excel: Optional[Path],
        summary: Dict[str, Any],
    ) -> Dict[str, str]:
        if supporting_excel is None:
            return {"status": "无截图文件", "detail": "未找到支持性附件截图.xlsx"}

        sheet_name = FULL_NODE_TO_SHEET.get(node, "")
        if not sheet_name:
            return {"status": "无支持性材料位置", "detail": f"节点{node}未配置支持性材料位置"}

        try:
            batch = self.evidence_loader.load_for_node(supporting_excel, node)
        except Exception as e:
            return {"status": "加载失败", "detail": str(e)}

        if batch is None or batch.image_count == 0:
            return {"status": "无截图", "detail": f"工作表'{sheet_name}'无嵌入图片"}

        # 需上传应用层账号截图: 只要截图存在即通过
        if node == "需上传应用层账号截图":
            return {
                "status": "通过",
                "detail": f"截图已上传({batch.image_count}张)",
                "image_count": str(batch.image_count),
            }

        # 截图文字识别 + 比对
        expected_str = summary.get("ExpectedAccounts", "")
        expected = [a.strip() for a in expected_str.split("|") if a.strip()]
        if not expected:
            return {"status": "无待核对账号", "detail": "账号审阅导出中无待核对账号"}

        try:
            pages = self.ocr_extractor.extract(batch)
            cr = self.evidence_checker.check(
                task_id=str(task.get("TaskID", "")),
                node=node,
                system=str(task.get("System", "")),
                expected_accounts=expected,
                ocr_pages=pages,
            )
        except Exception as e:
            return {"status": "截图核对异常", "detail": str(e)}

        passed = len(cr.passed)
        missing = len(cr.missing)
        global_confirmed = len(cr.spatial_mismatch)
        confirmed = passed
        total_confirmed = confirmed + global_confirmed
        matched_accounts = [m.expected_account for m in cr.passed]
        global_accounts = [m.expected_account for m in cr.spatial_mismatch]
        missing_accounts = [m.expected_account for m in cr.missing]
        common = {
            "image_count": str(batch.image_count),
            "ocr_recognized_accounts": " | ".join(cr.actual_accounts),
            "matched_accounts": " | ".join(matched_accounts),
            "global_matched_accounts": " | ".join(global_accounts),
            "missing_accounts": " | ".join(missing_accounts),
        }

        if cr.overall_status == "PASS":
            return {
                "status": "通过",
                "detail": f"已确认{confirmed}个，未确认0个",
                **common,
            }
        elif cr.overall_status == "REVIEW_SPATIAL":
            return {
                "status": "待确认",
                "detail": (
                    f"已确认{total_confirmed}个，未确认{missing}个；"
                    f"其中按表格位置确认{confirmed}个，全文补充确认{global_confirmed}个，"
                    "需人工确认全文补充结果"
                ),
                **common,
            }
        elif cr.overall_status == "REVIEW":
            return {
                "status": "待确认",
                "detail": f"已确认{confirmed}个，未确认{missing}个",
                **common,
            }
        else:
            return {
                "status": "未通过",
                "detail": f"已确认0个，未确认{missing}个",
                **common,
            }

    # ----------------------------------------------------------
    # 派生检查
    # ----------------------------------------------------------

    def _check_derived_nodes(
        self,
        node: str,
        folder_path: str,
        export_result: AccountExportReadResult,
        supporting_excel: Optional[Path],
        sid: str,
    ) -> Optional[Dict[str, str]]:
        if node not in {"需上传开发人员清单", "开发人员在生产环境无只读以上权限",
                         "无前后台管理员SOD问题", "应用层管理员角色完整性"}:
            return None

        folder = Path(folder_path)

        if node == "需上传开发人员清单":
            dev_list = self._resolve_dev_list(folder)
            exists = dev_list is not None
            return {
                "status": "通过" if exists else "未通过",
                "detail": "文件存在" if exists else "未找到开发人员清单.xlsx",
                "records": [{
                    "file_path": str(dev_list or ""),
                    "file_status": "已提供" if exists else "未提供",
                }],
            }

        if node == "开发人员在生产环境无只读以上权限":
            dev_list = self._resolve_dev_list(folder)
            if dev_list is None:
                return {"status": "未通过", "detail": "未找到开发人员清单.xlsx"}
            return self._check_dev_permissions(dev_list, folder, export_result)

        if node == "无前后台管理员SOD问题":
            return self._check_sod(export_result)

        if node == "应用层管理员角色完整性":
            return self._check_application_admin_role_completeness(sid, supporting_excel)

        return None

    def _check_application_admin_role_completeness(
        self, sid: str, supporting_excel: Optional[Path]
    ) -> Dict[str, Any]:
        """用 OCR 从“应用层管理员角色截图”的角色列核对 SID 配置角色。"""
        sid_key = normalize_sid(sid)
        if not sid_key:
            return {
                "status": "待确认",
                "detail": "WP 缺少 SID，无法读取应用层管理员角色配置。",
                "records": [],
            }

        rules = load_application_admin_role_rules().get(sid_key, [])
        if not rules:
            return {
                "status": "待确认",
                "detail": f"SID={sid} 未配置应用层管理员角色关键词。",
                "records": [],
            }
        if supporting_excel is None:
            return {
                "status": "未通过",
                "detail": "未找到支持性附件截图.xlsx，无法读取“应用层管理员角色截图”。",
                "records": [],
            }

        try:
            batch = self.evidence_loader.load_for_node(supporting_excel, "应用层管理员角色完整性")
        except Exception as exc:
            return {"status": "异常", "detail": f"加载角色截图失败：{exc}", "records": []}
        if batch is None or batch.image_count == 0:
            return {
                "status": "未通过",
                "detail": "工作表“应用层管理员角色截图”不存在或没有嵌入截图。",
                "records": [],
            }

        try:
            pages = self.ocr_extractor.extract(batch)
            actual_roles = self.evidence_checker.extract_column_values(
                pages, "应用层管理员角色完整性", "role_column"
            )
        except Exception as exc:
            return {"status": "异常", "detail": f"角色截图 OCR 失败：{exc}", "records": []}

        if not actual_roles:
            return {
                "status": "未通过",
                "detail": "OCR 未在截图中定位到“角色”列或该列没有可识别值。",
                "records": [],
            }

        records: List[Dict[str, str]] = []
        missing: List[str] = []
        for rule in rules:
            expected = rule["role_keyword"]
            match = None
            for actual in actual_roles:
                matched, method = match_role_keyword(expected, actual, rule["match_mode"])
                if matched:
                    match = (actual, method)
                    break
            if match:
                records.append({
                    "sid": sid,
                    "expected_role": expected,
                    "match_mode": "模糊" if rule["match_mode"] == "fuzzy" else "精确",
                    "actual_role": match[0],
                    "match_status": "一致",
                    "difference": "",
                    "match_method": match[1],
                })
            else:
                missing.append(expected)
                records.append({
                    "sid": sid,
                    "expected_role": expected,
                    "match_mode": "模糊" if rule["match_mode"] == "fuzzy" else "精确",
                    "actual_role": "",
                    "match_status": "不一致",
                    "difference": "截图角色列未识别到该关键词",
                    "match_method": "",
                })

        recognized = " | ".join(actual_roles)
        if missing:
            return {
                "status": "未通过",
                "detail": (
                    f"角色截图识别到：{recognized}；"
                    f"缺失配置角色：{' | '.join(missing)}"
                ),
                "records": records,
            }
        return {
            "status": "通过",
            "detail": f"SID={sid} 的配置角色均已覆盖；OCR 角色列：{recognized}",
            "records": records,
        }

    def _resolve_dev_list(self, folder: Path) -> Optional[Path]:
        """查找开发人员清单（兼容文件名编码损坏）。"""
        for required_columns in (["姓名"], ["开发人员姓名"], ["姓名（中文或英文全名）"], ["AD账号"], ["域账号"]):
            dev = self._find_file_by_content(
                folder, "开发人员清单.xlsx",
                required_columns=required_columns,
            )
            if dev is not None:
                return dev
        direct = folder / "开发人员清单.xlsx"
        return direct if direct.exists() else None

    def _check_dev_permissions(
        self, dev_list: Path, folder: Path, export_result: AccountExportReadResult
    ) -> Dict[str, str]:
        if not dev_list.exists():
            return {"status": "未通过", "detail": "未找到开发人员清单.xlsx"}
        if export_result.df is None:
            return {"status": "未通过", "detail": "无账号审阅导出数据"}

        try:
            dev_df = None
            for sheet_name, candidate in pd.read_excel(dev_list, sheet_name=None).items():
                candidate = candidate.rename(columns={
                    "姓名（中文或英文全名）": "姓名",
                    "开发人员姓名": "姓名",
                    "AD账号": "域账号",
                })
                name_col = resolve_column(candidate, ["姓名", "中文名", "英文名", "Name", "Developer", "开发人员"])
                acct_col = resolve_column(candidate, ["域账号", "AD账号", "AD Account", "ad_account", "账号"])
                if name_col and name_col != "姓名":
                    candidate["姓名"] = candidate[name_col]
                if acct_col and acct_col != "域账号":
                    candidate["域账号"] = candidate[acct_col]
                if "姓名" in candidate.columns or "域账号" in candidate.columns:
                    dev_df = candidate
                    break
            if dev_df is None:
                return {"status": "异常", "detail": "开发人员清单缺少姓名/域账号字段"}
            if "域账号" not in dev_df.columns:
                dev_df["域账号"] = ""
            if "姓名" not in dev_df.columns:
                dev_df["姓名"] = ""

            audit_df = export_result.df
            account_col = resolve_column(audit_df, EXPORT_ACCOUNT_CANDIDATES)
            owner_col = resolve_column(audit_df, EXPORT_OWNER_CANDIDATES)
            usage_col = resolve_column(audit_df, EXPORT_USAGE_CANDIDATES)
            if not account_col and not owner_col:
                return {"status": "异常", "detail": "账号审阅导出缺少账号名/账号使用人字段"}

            alias_map = self._load_name_alias_map()
            violations = []
            not_in_export = []
            developer_records = []
            matched_read_only = 0
            checked = 0
            for _, d in dev_df.iterrows():
                name = str(d.get("姓名", "")).strip()
                acct = str(d.get("域账号", "")).strip()
                if not name and not acct:
                    continue
                checked += 1
                matched = self._match_developer_in_export(
                    audit_df=audit_df,
                    account_col=account_col,
                    owner_col=owner_col,
                    name=name,
                    acct=acct,
                    alias_map=alias_map,
                )
                if matched.empty:
                    not_in_export.append(f"{name or '未提供姓名'}({acct or '无域账号'})")
                    developer_records.append({
                        "list_name": name,
                        "list_account": acct,
                        "export_account": "",
                        "export_owner": "",
                        "export_permission": "",
                        "match_status": "一致",
                        "difference": "未出现在审阅导出表，按规则通过",
                    })
                    continue
                for _, m in matched.iterrows():
                    usage = str(m.get(usage_col, "")).strip() if usage_col else ""
                    export_account = str(m.get(account_col, "")).strip() if account_col else ""
                    export_owner = str(m.get(owner_col, "")).strip() if owner_col else ""
                    if not self._is_read_only_usage(usage):
                        violations.append(f"{name}({acct}):{usage or '用途为空'}")
                        developer_records.append({
                            "list_name": name,
                            "list_account": acct,
                            "export_account": export_account,
                            "export_owner": export_owner,
                            "export_permission": usage,
                            "match_status": "不一致",
                            "difference": f"权限为{usage or '空'}，不是只读",
                        })
                    else:
                        matched_read_only += 1
                        developer_records.append({
                            "list_name": name,
                            "list_account": acct,
                            "export_account": export_account,
                            "export_owner": export_owner,
                            "export_permission": usage,
                            "match_status": "一致",
                            "difference": "",
                        })

            if violations:
                detail = (
                    f"开发人员在审阅导出中存在非只读权限({len(violations)}条): "
                    f"{'; '.join(violations[:5])}"
                )
                if not_in_export:
                    detail += (
                        f"；未出现在导出表按规则通过({len(not_in_export)}人): "
                        f"{'; '.join(not_in_export[:5])}"
                    )
                return {"status": "未通过", "detail": detail, "records": developer_records}
            if checked == 0:
                return {"status": "异常", "detail": "开发人员清单无有效人员记录"}
            detail_parts = []
            if matched_read_only:
                detail_parts.append(f"导出表中匹配到的开发人员均为只读({matched_read_only}条)")
            if not_in_export:
                detail_parts.append(
                    f"未出现在审阅导出表按规则通过({len(not_in_export)}人): "
                    f"{'; '.join(not_in_export[:5])}"
                )
                detail_parts.append("如确认为中英名未匹配，请维护 configs/中英文姓名对照表.xlsx")
            return {
                "status": "通过",
                "detail": "；".join(detail_parts) or f"开发人员权限检查通过({checked}人)",
                "records": developer_records,
            }
        except Exception as e:
            return {"status": "异常", "detail": str(e)}

    @staticmethod
    def _is_read_only_usage(value: str) -> bool:
        text = str(value).strip().upper().replace(" ", "").replace("-", "").replace("_", "")
        return text in {"只读", "READONLY"}

    @staticmethod
    def _add_alias_group(alias_map: Dict[str, set], values: List[str]) -> None:
        aliases = {v.strip() for v in values if str(v).strip()}
        if not aliases:
            return
        keys = {normalize_key(v) for v in aliases if normalize_key(v)}
        for key in keys:
            alias_map.setdefault(key, set()).update(aliases)

    def _load_name_alias_map(self) -> Dict[str, set]:
        """
        读取开发人员中英名/账号映射。

        优先支持 configs/中英文姓名对照表.xlsx；同时兼容旧英文文件名和账号别名表。
        """
        alias_map: Dict[str, set] = {}
        for mapping_file in (
            NAME_ALIAS_FILE,
            LEGACY_NAME_ALIAS_FILE,
            ACCOUNT_ALIAS_FILE,
            LEGACY_ACCOUNT_ALIAS_FILE,
        ):
            if not mapping_file.exists():
                continue
            try:
                sheets = pd.read_excel(mapping_file, sheet_name=None)
            except Exception:
                continue
            for _, df in sheets.items():
                if "enabled" in df.columns:
                    disabled = {"FALSE", "F", "N", "NO", "否", "0", "禁用"}
                    enabled_text = df["enabled"].astype(str).str.strip().str.upper()
                    df = df[~enabled_text.isin(disabled)]
                for _, row in df.iterrows():
                    values: List[str] = []
                    for col in (
                        "standard_name", "alias", "english_name", "chinese_name",
                        "ad_account", "raw_name", "standard_account", "姓名", "英文名", "中文名", "域账号",
                        "原始写法", "标准账号", "其他别名", "AD账号",
                    ):
                        if col not in df.columns:
                            continue
                        values.extend(split_multi_values(row.get(col, "")))
                    self._add_alias_group(alias_map, values)
        return alias_map

    def _expand_developer_terms(self, name: str, acct: str, alias_map: Dict[str, set]) -> List[str]:
        terms = [name, acct]
        for raw in (name, acct):
            key = normalize_key(raw)
            if key and key in alias_map:
                terms.extend(alias_map[key])
        seen = set()
        result = []
        for term in terms:
            key = normalize_key(term)
            if key and key not in seen:
                seen.add(key)
                result.append(str(term).strip())
        return result

    @staticmethod
    def _owner_matches_term(owner: str, term: str) -> bool:
        owner_text = str(owner).strip()
        term_text = str(term).strip()
        owner_key = normalize_key(owner_text)
        term_key = normalize_key(term_text)
        if not owner_key or not term_key:
            return False
        if owner_key == term_key:
            return True
        for part in split_multi_values(owner_text):
            if normalize_key(part) == term_key:
                return True
        has_cjk = any("\u4e00" <= ch <= "\u9fff" for ch in term_text)
        if has_cjk and len(term_text) >= 2 and term_text in owner_text:
            return True
        if not has_cjk and len(term_key) >= 4 and term_key in owner_key:
            return True
        return False

    def _match_developer_in_export(
        self,
        audit_df: pd.DataFrame,
        account_col: Optional[str],
        owner_col: Optional[str],
        name: str,
        acct: str,
        alias_map: Dict[str, set],
    ) -> pd.DataFrame:
        terms = self._expand_developer_terms(name, acct, alias_map)
        matched_indices = set()

        if account_col:
            account_keys = audit_df[account_col].astype(str).map(normalize_key)
            term_keys = {normalize_key(t) for t in terms if normalize_key(t)}
            if term_keys:
                matched_indices.update(audit_df[account_keys.isin(term_keys)].index)

        if owner_col:
            for idx, owner in audit_df[owner_col].items():
                if any(self._owner_matches_term(str(owner), term) for term in terms):
                    matched_indices.add(idx)

        if not matched_indices:
            return audit_df.iloc[0:0].copy()
        return audit_df.loc[sorted(matched_indices)].copy()

    @staticmethod
    def _check_sod(export_result: AccountExportReadResult) -> Dict[str, str]:
        if export_result.df is None:
            return {"status": "未通过", "detail": "无账号审阅导出数据"}

        df = export_result.df
        APP = {"应用层"}
        DB = {"数据库", "应用服务器", "发布工具服务器", "数据库服务器"}
        ADMIN = {"管理员", "读写"}
        SKIP = {"系统", "系统账号", "PAM", "",
                "登录服务器后的读写账号", "登录服务器后的只读账号", "系统默认账号"}

        app_df = df[df["审阅对象"].isin(APP) & df["用途"].isin(ADMIN)].copy()
        db_df = df[df["审阅对象"].isin(DB) & df["用途"].isin(ADMIN)].copy()
        app_users = set(app_df["账号使用人"].dropna().astype(str).str.strip())
        db_users = set(db_df["账号使用人"].dropna().astype(str).str.strip())
        overlaps = (app_users & db_users) - SKIP

        def summarize_rows(part: pd.DataFrame) -> str:
            values = []
            for _, row in part.iterrows():
                acct = str(row.get("账号名", "")).strip()
                usage = str(row.get("用途", "")).strip()
                obj = str(row.get("审阅对象", "")).strip()
                if acct or usage or obj:
                    values.append(f"{acct}({obj}/{usage})")
            return "；".join(values)

        records = []
        if overlaps:
            for owner in sorted(overlaps):
                records.append({
                    "person": owner,
                    "app_accounts": summarize_rows(app_df[app_df["账号使用人"].astype(str).str.strip() == owner]),
                    "backend_accounts": summarize_rows(db_df[db_df["账号使用人"].astype(str).str.strip() == owner]),
                    "match_status": "不一致",
                    "difference": "前后台管理员存在同一人员",
                })
        else:
            records.append({
                "person": "无",
                "app_accounts": summarize_rows(app_df),
                "backend_accounts": summarize_rows(db_df),
                "match_status": "一致",
                "difference": "无前后台管理员重合",
            })

        if overlaps:
            return {"status": "未通过",
                    "detail": f"前后台管理员重合: {', '.join(sorted(overlaps))}",
                    "records": records}
        return {"status": "通过", "detail": "无前后台管理员重合", "records": records}

    # ----------------------------------------------------------
    # 汇总最终结论
    # ----------------------------------------------------------

    @staticmethod
    def _compute_final(
        account_ok: Optional[bool],
        account_detail: str,
        manual_overall: str,
        parse_method: str,
        ocr_needed: bool,
        ocr_ok: Optional[bool],
        ocr_status: str,
        ocr_detail: str,
        derived_result: Optional[Dict[str, str]],
    ) -> tuple[str, str, str]:
        # 派生检查优先
        if derived_result:
            s = derived_result.get("status", "")
            detail = derived_result.get("detail", s)
            if s == "通过":
                return "是", detail, "通过"
            if s == "异常":
                return f"异常-{detail}", detail, "异常"
            if s == "待确认":
                return f"待确认-{detail}", detail, "待确认"
            return f"否-{detail}", detail, "未通过"

        # 账号不存在 → 失败
        if account_ok is False:
            detail = account_detail or "账号检查未通过"
            return f"否-{detail}", detail, "未通过"

        # 需要 OCR
        if ocr_needed:
            detail = ocr_detail or ocr_status or "截图待检查"
            if ocr_ok is True:
                return "是", detail, "通过"
            if ocr_status in {"截图核对异常", "加载失败"}:
                return f"异常-{detail}", detail, "异常"
            if ocr_status in {"待确认", "部分通过"}:
                return f"待确认-{detail}", detail, "待确认"
            return f"否-{detail}", detail, "未通过"

        # 不需 OCR 只能由明确的业务路径通过，未知/混合状态不得自动通过。
        if account_ok is True and parse_method == "no_manual_upload" and manual_overall == "N":
            detail = "支持性材料标识=N，该节点无需上传截图支持文件"
            return "是-无需上传截图支持文件", detail, "通过"
        if account_ok is True and parse_method == "export_available":
            detail = "已通过账号审阅导出验证，该节点无需上传截图支持文件"
            return "是-系统导出验证", detail, "通过"

        detail = (
            f"未满足明确通过条件：检查方式={parse_method or '未配置'}, "
            f"支持性材料标识={manual_overall or '待确认'}"
        )
        return f"待确认-{detail}", detail, "待确认"

    # ----------------------------------------------------------
    # 构建最终 Excel
    # ----------------------------------------------------------

    @staticmethod
    def _split_accounts(value: Any) -> List[str]:
        return [x.strip() for x in str(value or "").split("|") if x.strip()]

    @staticmethod
    def _sid_from_wp_row(wp_row: pd.Series) -> str:
        """SID 只接受明确列名，不回退到系统名称。"""
        for column in ("SID", "系统ID"):
            if column in wp_row.index:
                return normalize_sid(wp_row.get(column, ""))
        return ""

    @staticmethod
    def _test_requirement_label(scope_label: str) -> str:
        text = str(scope_label or "").strip()
        if text.startswith("N/A"):
            return "不测试"
        if text.startswith("Refer to"):
            return "引用其他材料"
        if text == "待确认":
            return "待确认"
        if text:
            return "需要测试"
        return "未评估"

    @staticmethod
    def _business_result_label(result: Dict[str, Any]) -> str:
        status = str(result.get("final_status", "")).strip()
        final = str(result.get("final", "")).strip()
        if status:
            return status
        if final.startswith("是"):
            return "通过"
        if final.startswith("否"):
            return "未通过"
        if final.startswith("N/A"):
            return "不适用"
        if final.startswith("Refer to"):
            return "引用材料"
        return "待确认"

    def _how_to_test(self, node: str, result: Dict[str, Any]) -> str:
        scope = str(result.get("scope", "")).strip()
        if scope.startswith("N/A"):
            return "不需要测试"
        if scope.startswith("Refer to"):
            return f"引用其他系统材料：{scope.removeprefix('Refer to').strip()}"
        if node == "需上传开发人员清单":
            return "检查开发人员清单是否已提供且可读取"
        if node == "开发人员在生产环境无只读以上权限":
            return "核对开发人员清单与账号审阅导出，确认匹配到的开发人员均为只读权限；未出现在导出表按规则通过"
        if node == "无前后台管理员SOD问题":
            return "核对应用层管理员与后台管理员是否存在同一人员重合"
        if node == "应用层管理员角色完整性":
            return "从“应用层管理员角色截图”的角色列 OCR 识别结果，核对是否覆盖 SID 配置的角色关键词"
        if result.get("ocr_needed") == "是":
            return "将账号审阅导出中的账号与截图识别结果逐项核对"
        parse_method = str(result.get("parse_method", "")).strip()
        if parse_method == "export_available":
            return "检查账号审阅导出中是否存在对应账号"
        if parse_method == "no_manual_upload":
            return "按账号审阅导出和支持性材料标识判断，无需截图核对"
        if result.get("manual_upload") == "Y":
            return "根据已上传支持性材料进行核对"
        return "按系统适用规则进行检查"

    def _build_test_scope_df(
        self, wp_df: pd.DataFrame, system_results: Dict[str, Dict[str, Dict]]
    ) -> pd.DataFrame:
        rows = []
        for _, wp_row in wp_df.iterrows():
            system = str(wp_row.get("System", "")).strip()
            sid = self._sid_from_wp_row(wp_row)
            node_results = system_results.get(sid, {})
            for node in AuditNodes.ALL:
                result = node_results.get(node, {})
                scope = str(result.get("scope", "未评估"))
                rows.append({
                    "Excel行号": int(wp_row.name) + 2 if isinstance(wp_row.name, int) else "",
                    "系统": system,
                    "SID": sid,
                    "系统类别": wp_row.get("Category", ""),
                    "检查节点": node,
                    "是否需要测试": self._test_requirement_label(scope),
                    "怎么测": self._how_to_test(node, result),
                    "判断依据": result.get("scope_reason", "") or result.get("scope_warning", ""),
                })
        return pd.DataFrame(rows)

    def _build_test_result_df(
        self, wp_df: pd.DataFrame, system_results: Dict[str, Dict[str, Dict]]
    ) -> pd.DataFrame:
        rows: List[Dict[str, Any]] = []
        for _, wp_row in wp_df.iterrows():
            system = str(wp_row.get("System", "")).strip()
            sid = self._sid_from_wp_row(wp_row)
            node_results = system_results.get(sid, {})
            for node in AuditNodes.ALL:
                result = node_results.get(node, {})
                scope = str(result.get("scope", "未评估"))
                rows.append({
                    "Excel行号": int(wp_row.name) + 2 if isinstance(wp_row.name, int) else "",
                    "系统": system,
                    "SID": sid,
                    "系统类别": wp_row.get("Category", ""),
                    "检查节点": node,
                    "是否需要测试": self._test_requirement_label(scope),
                    "怎么测": self._how_to_test(node, result),
                    "测试结果": self._business_result_label(result),
                    "结论": result.get("final", "未评估"),
                    "结果说明": result.get("final_reason", ""),
                })
        return pd.DataFrame(rows)

    def _build_business_detail_sheets(
        self, wp_df: pd.DataFrame, system_results: Dict[str, Dict[str, Dict]]
    ) -> Dict[str, pd.DataFrame]:
        """按业务复核类型拆分明细，避免单张明细表出现大量空列。"""
        screenshot_rows: List[Dict[str, Any]] = []
        export_rows: List[Dict[str, Any]] = []
        developer_rows: List[Dict[str, Any]] = []
        sod_rows: List[Dict[str, Any]] = []
        app_role_rows: List[Dict[str, Any]] = []

        def base(wp_row: pd.Series, system: str, node: str, result: Dict[str, Any]) -> Dict[str, Any]:
            return {
                "Excel行号": int(wp_row.name) + 2 if isinstance(wp_row.name, int) else "",
                "系统": system,
                "SID": self._sid_from_wp_row(wp_row),
                "系统类别": wp_row.get("Category", ""),
                "检查节点": node,
                "测试结果": self._business_result_label(result),
                "说明": result.get("final_reason", ""),
            }

        for _, wp_row in wp_df.iterrows():
            system = str(wp_row.get("System", "")).strip()
            sid = self._sid_from_wp_row(wp_row)
            node_results = system_results.get(sid, {})
            for node in AuditNodes.ALL:
                result = node_results.get(node, {})
                scope = str(result.get("scope", "")).strip()
                if scope.startswith("N/A") or scope.startswith("Refer to"):
                    continue

                expected_accounts = self._split_accounts(result.get("expected_accounts", ""))
                structure_accounts = set(self._split_accounts(result.get("ocr_matched_accounts", "")))
                full_text_accounts = set(self._split_accounts(result.get("ocr_global_matched_accounts", "")))
                missing_accounts = set(self._split_accounts(result.get("ocr_missing_accounts", "")))

                # 1) 截图核对：只展示从账号审阅导出出发、逐项到截图中确认的账号。
                if result.get("ocr_needed") == "是" or structure_accounts or full_text_accounts or missing_accounts:
                    if (
                        result.get("ocr_needed") == "是"
                        and self._business_result_label(result) == "通过"
                        and expected_accounts
                        and not structure_accounts
                        and not full_text_accounts
                        and not missing_accounts
                    ):
                        structure_accounts = set(expected_accounts)

                    for account in expected_accounts:
                        row = base(wp_row, system, node, result)
                        row["系统提取出的账号"] = account
                        if account in structure_accounts:
                            row.update({
                                "截图识别出的账号": account,
                                "核对方式": "按表格位置确认",
                                "是否一致": "一致",
                                "哪里不一致": "",
                            })
                        elif account in full_text_accounts:
                            row.update({
                                "截图识别出的账号": account,
                                "核对方式": "全文补充确认",
                                "是否一致": "待确认",
                                "哪里不一致": "账号值一致，但仅在截图全文中找到，未在表格位置确认",
                            })
                        elif account in missing_accounts:
                            row.update({
                                "截图识别出的账号": "",
                                "核对方式": "未确认",
                                "是否一致": "不一致",
                                "哪里不一致": "截图中未确认到该账号",
                            })
                        else:
                            row.update({
                                "截图识别出的账号": "",
                                "核对方式": "未形成明确结果",
                                "是否一致": "待确认",
                                "哪里不一致": "未形成明确比对结果",
                            })
                        screenshot_rows.append(row)
                    continue

                # 2) 账号审阅导出存在性核对：只展示导出中是否存在、有哪些账号。
                if result.get("parse_method") == "export_available":
                    row = base(wp_row, system, node, result)
                    row.update({
                        "审阅对象": result.get("review_object", ""),
                        "是否存在账号": "是" if expected_accounts else "否",
                        "账号数量": len(expected_accounts),
                        "账号清单": " | ".join(expected_accounts) if expected_accounts else "",
                    })
                    export_rows.append(row)
                    continue

                # 3) 开发人员：清单提供情况 + 清单和导出权限对比。
                if node == "需上传开发人员清单":
                    records = result.get("derived_records", []) or []
                    for record in records:
                        row = base(wp_row, system, node, result)
                        row.update({
                            "清单路径": record.get("file_path", ""),
                            "清单状态": record.get("file_status", ""),
                            "清单姓名": "",
                            "清单账号": "",
                            "导出账号": "",
                            "导出使用人": "",
                            "导出权限": "",
                            "是否一致": "一致" if record.get("file_status") == "已提供" else "不一致",
                            "哪里不一致": "" if record.get("file_status") == "已提供" else "未找到开发人员清单",
                        })
                        developer_rows.append(row)
                    continue

                if node == "开发人员在生产环境无只读以上权限":
                    records = result.get("derived_records", []) or []
                    for record in records:
                        row = base(wp_row, system, node, result)
                        row.update({
                            "清单路径": "",
                            "清单状态": "",
                            "清单姓名": record.get("list_name", ""),
                            "清单账号": record.get("list_account", ""),
                            "导出账号": record.get("export_account", ""),
                            "导出使用人": record.get("export_owner", ""),
                            "导出权限": record.get("export_permission", ""),
                            "是否一致": record.get("match_status", ""),
                            "哪里不一致": record.get("difference", ""),
                        })
                        developer_rows.append(row)
                    continue

                # 4) SOD：展示应用层管理员和后台管理员名单/权限；没有重合则写“无”。
                if node == "无前后台管理员SOD问题":
                    records = result.get("derived_records", []) or []
                    for record in records:
                        row = base(wp_row, system, node, result)
                        row.update({
                            "人员": record.get("person", ""),
                            "应用层管理员账号及权限": record.get("app_accounts", ""),
                            "后台管理员账号及权限": record.get("backend_accounts", ""),
                            "是否一致": record.get("match_status", ""),
                            "哪里不一致": record.get("difference", ""),
                        })
                        sod_rows.append(row)
                    continue

                # 5) 应用层管理员角色完整性：逐个展示 SID 配置角色与 OCR 角色列的比对。
                if node == "应用层管理员角色完整性":
                    records = result.get("derived_records", []) or []
                    for record in records:
                        row = base(wp_row, system, node, result)
                        row.update({
                            "配置角色关键词": record.get("expected_role", ""),
                            "匹配方式": record.get("match_mode", ""),
                            "截图识别角色": record.get("actual_role", ""),
                            "匹配结果": record.get("match_status", ""),
                            "匹配说明": record.get("difference", "") or record.get("match_method", ""),
                        })
                        app_role_rows.append(row)
                    continue

        return {
            "明细-截图核对": pd.DataFrame(
                screenshot_rows,
                columns=[
                    "Excel行号", "系统", "SID", "系统类别", "检查节点", "系统提取出的账号",
                    "截图识别出的账号", "核对方式", "是否一致", "哪里不一致", "测试结果", "说明",
                ],
            ),
            "明细-导出核对": pd.DataFrame(
                export_rows,
                columns=[
                    "Excel行号", "系统", "SID", "系统类别", "检查节点", "审阅对象",
                    "是否存在账号", "账号数量", "账号清单", "测试结果", "说明",
                ],
            ),
            "明细-开发人员": pd.DataFrame(
                developer_rows,
                columns=[
                    "Excel行号", "系统", "SID", "系统类别", "检查节点", "清单路径", "清单状态",
                    "清单姓名", "清单账号", "导出账号", "导出使用人", "导出权限",
                    "是否一致", "哪里不一致", "测试结果", "说明",
                ],
            ),
            "明细-SOD": pd.DataFrame(
                sod_rows,
                columns=[
                    "Excel行号", "系统", "SID", "系统类别", "检查节点", "人员",
                    "应用层管理员账号及权限", "后台管理员账号及权限",
                    "是否一致", "哪里不一致", "测试结果", "说明",
                ],
            ),
            "明细-应用层管理员角色": pd.DataFrame(
                app_role_rows,
                columns=[
                    "Excel行号", "系统", "SID", "系统类别", "检查节点", "配置角色关键词",
                    "匹配方式", "截图识别角色", "匹配结果", "匹配说明", "测试结果", "说明",
                ],
            ),
        }

    def _build_test_detail_df(
        self, wp_df: pd.DataFrame, system_results: Dict[str, Dict[str, Dict]]
    ) -> pd.DataFrame:
        """业务可读的明细：重点展示账号差异、截图比对差异、开发人员权限差异。"""
        rows: List[Dict[str, Any]] = []

        def add_row(wp_row: pd.Series, system: str, node: str, result: Dict[str, Any], **extra: Any) -> None:
            base = {
                "Excel行号": int(wp_row.name) + 2 if isinstance(wp_row.name, int) else "",
                "系统": system,
                "系统类别": wp_row.get("Category", ""),
                "检查节点": node,
                "是否需要测试": self._test_requirement_label(str(result.get("scope", ""))),
                "怎么测": self._how_to_test(node, result),
                "测试结果": self._business_result_label(result),
                "系统提取出的账号": "",
                "截图识别出的账号": "",
                "是否一致": "",
                "哪里不一致": "",
                "清单姓名": "",
                "清单账号": "",
                "导出账号": "",
                "导出使用人": "",
                "导出权限": "",
                "说明": result.get("final_reason", ""),
            }
            base.update(extra)
            rows.append(base)

        for _, wp_row in wp_df.iterrows():
            system = str(wp_row.get("System", "")).strip()
            sid = self._sid_from_wp_row(wp_row)
            node_results = system_results.get(sid, {})
            for node in AuditNodes.ALL:
                result = node_results.get(node, {})

                developer_records = result.get("derived_records", []) or []
                if node == "开发人员在生产环境无只读以上权限" and developer_records:
                    for record in developer_records:
                        add_row(
                            wp_row, system, node, result,
                            清单姓名=record.get("list_name", ""),
                            清单账号=record.get("list_account", ""),
                            导出账号=record.get("export_account", ""),
                            导出使用人=record.get("export_owner", ""),
                            导出权限=record.get("export_permission", ""),
                            是否一致=record.get("match_status", ""),
                            哪里不一致=record.get("difference", ""),
                            说明=result.get("final_reason", ""),
                        )
                    continue

                expected_accounts = self._split_accounts(result.get("expected_accounts", ""))
                structure_accounts = set(self._split_accounts(result.get("ocr_matched_accounts", "")))
                full_text_accounts = set(self._split_accounts(result.get("ocr_global_matched_accounts", "")))
                missing_accounts = set(self._split_accounts(result.get("ocr_missing_accounts", "")))
                if (
                    result.get("ocr_needed") == "是"
                    and self._business_result_label(result) == "通过"
                    and expected_accounts
                    and not structure_accounts
                    and not full_text_accounts
                    and not missing_accounts
                ):
                    structure_accounts = set(expected_accounts)

                if result.get("ocr_needed") == "是" or structure_accounts or full_text_accounts or missing_accounts:
                    for account in expected_accounts:
                        if account in structure_accounts:
                            add_row(
                                wp_row, system, node, result,
                                系统提取出的账号=account,
                                截图识别出的账号=account,
                                是否一致="一致",
                                哪里不一致="",
                                说明="按表格位置确认一致",
                            )
                        elif account in full_text_accounts:
                            add_row(
                                wp_row, system, node, result,
                                系统提取出的账号=account,
                                截图识别出的账号=account,
                                是否一致="待确认",
                                哪里不一致="账号值一致，但仅在截图全文中找到，未在表格位置确认",
                                说明="需要人工确认全文补充结果",
                            )
                        elif account in missing_accounts:
                            add_row(
                                wp_row, system, node, result,
                                系统提取出的账号=account,
                                截图识别出的账号="",
                                是否一致="不一致",
                                哪里不一致="截图中未确认到该账号",
                                说明=result.get("final_reason", ""),
                            )
                        else:
                            add_row(
                                wp_row, system, node, result,
                                系统提取出的账号=account,
                                截图识别出的账号="",
                                是否一致="待确认",
                                哪里不一致="未形成明确比对结果",
                                说明=result.get("final_reason", ""),
                            )
                    continue

                add_row(
                    wp_row, system, node, result,
                    系统提取出的账号=result.get("expected_accounts", ""),
                    截图识别出的账号="不涉及截图核对",
                    是否一致="不适用",
                    哪里不一致="" if self._business_result_label(result) == "通过" else result.get("final_reason", ""),
                )
        return pd.DataFrame(rows)


# ============================================================
# CLI
# ============================================================

def main() -> None:
    parser = argparse.ArgumentParser(description="账号审阅端到端流水线")
    parser.add_argument("--input", default=str(DEFAULT_INPUT), help="WP Excel 路径")
    parser.add_argument("--sheet", default="Sheet1", help="Sheet 名")
    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"❌ 文件不存在: {input_path}")
        return

    wp_df = AccountAuditPipeline.read_file(str(input_path), sheet_name=args.sheet)
    print(f"读取 WP: {input_path} [{len(wp_df)} 行]")

    result = AuditConclusionGenerator().generate(wp_df)
    print(f"\n✅ 审阅结论: {result['conclusion']}")


if __name__ == "__main__":
    main()
