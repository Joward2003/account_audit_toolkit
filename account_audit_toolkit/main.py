"""Run account audit workflow export — 完整版。

端到端流水线，将 scope engine、execution plan builder、evidence checker 串联，
生成一份与 WP 模板（test.xlsx）对齐的审阅结论 Excel。

流水线:
  WP Excel (test.xlsx)
    → Scope Engine (哪些节点要查)
    → System-Folder Mapper (映射到 supporting 文件夹)
    → Execution Plan Builder (build_expected_accounts + decide_evidence_plan)
    → Evidence Checker (OCR 截图比对，仅 手工上传标识=Y 时)
    → 最终审阅结论 Excel

输出:
  output/account_audit_conclusion.xlsx
    - Sheet「审阅结论」: WP 原列 + 10 个审计节点结论 + 通过率
    - Sheet「Scope明细」: Pipeline 完整明细

用法:
  python main.py
  python main.py --input ../test.xlsx --sheet Sheet1
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any, Dict, List, Optional

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
)

# ---- 直接复用 evidence_checker ----
from evidence_checker import (
    EvidenceLoader,
    OCRExtractor,
    EvidenceChecker,
    FULL_NODE_TO_SHEET,
)


# ============================================================
# 配置
# ============================================================

BASE_DIR = Path(__file__).resolve().parent.parent
SUPPORTING_DIR = BASE_DIR / "supporting"
OUTPUT_DIR = BASE_DIR / "account_audit_toolkit" / "output"
DEFAULT_INPUT = BASE_DIR / "test.xlsx"
CONCLUSION_FILE = OUTPUT_DIR / "account_audit_conclusion.xlsx"


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
        self, wp_df: pd.DataFrame, supporting_dir: Optional[Path] = None
    ) -> Dict[str, Any]:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        _supporting = supporting_dir or SUPPORTING_DIR

        # Step 1: Scope Engine
        print("=" * 60)
        print("Step 1: Scope Engine")
        print("=" * 60)
        scope_results = self.exporter.build_scope_results(wp_df)
        todo_df = self.exporter.build_todo_df(scope_results)
        detail_df = self.pipeline.run(wp_df)
        for _, row in wp_df.iterrows():
            print(f"  {row['System']:15s} | {row['Category']}")

        # Step 2: Folder mapping
        print("\n" + "=" * 60)
        print("Step 2: Folder Mapper")
        print("=" * 60)
        folder_map = self._build_folder_map(_supporting)
        if supporting_dir:
            print(f"  [使用上传的 supporting: {_supporting}]", flush=True)
            print(f"  [文件夹映射: { {k: Path(v).name for k, v in folder_map.items()} }]", flush=True)
        for sys_name in pd.unique(todo_df["System"]):
            fp = self._resolve_folder(str(sys_name), folder_map)
            print(f"  {str(sys_name):15s} → {fp or 'NOT FOUND'}")

        # Step 3: Per-system check
        print("\n" + "=" * 60)
        print("Step 3: Evidence Check（账号审阅导出 + 截图OCR仅Y时）")
        print("=" * 60)

        system_results: Dict[str, Dict[str, Dict]] = {}
        for system in pd.unique(todo_df["System"]):
            system_str = str(system).strip()
            folder_path = self._resolve_folder(system_str, folder_map)
            system_results[system_str] = self._check_system(
                system_str, folder_path, todo_df
            )

        # Step 4: Build conclusion Excel
        print("\n" + "=" * 60)
        print("Step 4: 生成审阅结论 Excel")
        print("=" * 60)
        conclusion_df = self._build_conclusion_df(wp_df, system_results)

        with pd.ExcelWriter(CONCLUSION_FILE, engine="openpyxl") as writer:
            conclusion_df.to_excel(writer, sheet_name="审阅结论", index=False)
            detail_df.to_excel(writer, sheet_name="Scope明细", index=False)

        print(f"  ✅ {CONCLUSION_FILE}")
        return {"conclusion": str(CONCLUSION_FILE)}

    # ----------------------------------------------------------
    # Folder mapping
    # ----------------------------------------------------------

    @staticmethod
    def _build_folder_map(root_dir: Path) -> Dict[str, str]:
        m: Dict[str, str] = {}
        for d in root_dir.iterdir():
            if not d.is_dir():
                continue
            if d.name.startswith(".") or d.name.startswith("._") or d.name == "__MACOSX":
                continue
            key = d.name.upper().replace(" ", "").replace("_", "").replace("-", "")
            m[key] = str(d)
        return m

    @staticmethod
    def _resolve_folder(system: str, folder_map: Dict[str, str]) -> str:
        key = system.upper().replace(" ", "").replace("_", "").replace("-", "")
        if key in folder_map:
            return folder_map[key]
        for k, v in folder_map.items():
            if key.startswith(k) or k.startswith(key):
                return v
        return ""

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
        self, system: str, folder_path: str, todo_df: pd.DataFrame
    ) -> Dict[str, Dict]:
        print(f"\n--- {system} ---")
        print(f"    folder_path={folder_path}", flush=True)

        if not folder_path:
            print("  ⚠ 无文件夹映射")
            return {}

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

        # 获取该系统任务
        system_todo = todo_df[todo_df["System"] == system]

        conclusions: Dict[str, Dict] = {}
        for _, task in system_todo.iterrows():
            node = str(task.get("检查节点", ""))
            scope_label = str(task.get("Scope结果", ""))
            conclusions[node] = self._check_one_node(
                node, scope_label, task, export_result, folder_path, supporting_excel
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
    ) -> Dict[str, Any]:
        """对单个审计节点执行完整测试逻辑。"""

        # Scope 判定：N/A 直接返回
        scope_is_yes = scope_label in {"是", "Yes"} or str(scope_label).startswith("Refer to")
        if not scope_is_yes:
            result = {
                "scope": scope_label,
                "account_ok": "N/A",
                "manual_upload": "N/A",
                "ocr": "N/A",
                "final": scope_label if scope_label else "N/A",
            }
            print(f"  [{node}] scope={scope_label} → {result['final']}")
            return result

        # ---- 1. 调用 execution_plan_builder 构建 expected accounts ----
        review_object = NODE_TO_REVIEW_OBJECT.get(node, "")
        if node in EVIDENCE_ONLY_NODE_REVIEW_OBJECT:
            review_object = EVIDENCE_ONLY_NODE_REVIEW_OBJECT[node]
        _, summary = build_expected_accounts_for_task(export_result, task, review_object)

        expected_count = summary.get("ExpectedCount", 0)
        expected_status = summary.get("ExpectedSetStatus", "")
        manual_overall = str(summary.get("ManualUploadOverall", "UNKNOWN")).strip()

        # ---- 2. 调用 execution_plan_builder 判断证据计划 ----
        plan = decide_evidence_plan(task, summary)
        parse_method = plan.get("ParseMethod", "")

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
        ocr_msg = "无需OCR"

        if ocr_needed:
            ocr_result = self._run_ocr_check(node, task, supporting_excel, summary)
            ocr_ok = ocr_result.get("status", "") == "通过"
            ocr_msg = ocr_result.get("status", "OCR失败")

        # ---- 5. 派生检查（开发人员清单/SOD） ----
        derived_result = self._check_derived_nodes(node, folder_path, export_result)

        # ---- 6. 汇总 ----
        final = self._compute_final(
            account_ok=account_ok,
            ocr_needed=ocr_needed,
            ocr_ok=ocr_ok,
            derived_result=derived_result,
        )

        print(f"  [{node}] account={account_msg} | upload={manual_overall} | "
              f"parse={parse_method} | ocr={'Y' if ocr_needed else 'N'}={ocr_msg}"
              f"{' | derived=' + derived_result.get('status','?') if derived_result else ''}"
              f" | → {final}")

        return {
            "scope": scope_label,
            "account_ok": account_msg,
            "manual_upload": manual_overall,
            "ocr": ocr_msg,
            "parse_method": parse_method,
            "final": final,
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
            return {"status": "无Sheet映射", "detail": f"节点{node}无supporting sheet"}

        try:
            batch = self.evidence_loader.load_for_node(supporting_excel, node)
        except Exception as e:
            return {"status": "加载失败", "detail": str(e)}

        if batch is None or batch.image_count == 0:
            return {"status": "无截图", "detail": f"Sheet'{sheet_name}'无嵌入图片"}

        # 需上传应用层账号截图: 只要截图存在即通过
        if node == "需上传应用层账号截图":
            return {"status": "通过", "detail": f"截图已上传({batch.image_count}张)"}

        # OCR + 比对
        expected_str = summary.get("ExpectedAccounts", "")
        expected = [a.strip() for a in expected_str.split("|") if a.strip()]
        if not expected:
            return {"status": "无预期账号", "detail": "账号导出无expected set"}

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
            return {"status": "OCR异常", "detail": str(e)}

        passed = len(cr.passed)
        missing = len(cr.missing)
        spatial = len(cr.spatial_mismatch)

        if cr.overall_status == "PASS":
            return {"status": "通过", "detail": f"全部匹配({passed}/{len(expected)})"}
        elif cr.overall_status == "REVIEW_SPATIAL":
            return {"status": "兜底匹配", "detail": f"锚定={passed},兜底={spatial},缺失={missing}"}
        elif cr.overall_status == "REVIEW":
            return {"status": "部分通过", "detail": f"匹配={passed},缺失={missing}"}
        else:
            return {"status": "未通过", "detail": f"缺失={missing}/{len(expected)}"}

    # ----------------------------------------------------------
    # 派生检查
    # ----------------------------------------------------------

    def _check_derived_nodes(
        self, node: str, folder_path: str, export_result: AccountExportReadResult
    ) -> Optional[Dict[str, str]]:
        if node not in {"需上传开发人员清单", "开发人员在生产环境无只读以上权限",
                         "无前后台管理员SOD问题"}:
            return None

        folder = Path(folder_path)

        if node == "需上传开发人员清单":
            dev_list = self._resolve_dev_list(folder)
            exists = dev_list is not None
            return {"status": "通过" if exists else "未通过",
                    "detail": "文件存在" if exists else "未找到开发人员清单.xlsx"}

        if node == "开发人员在生产环境无只读以上权限":
            dev_list = self._resolve_dev_list(folder)
            if dev_list is None:
                return {"status": "未通过", "detail": "未找到开发人员清单.xlsx"}
            return self._check_dev_permissions(dev_list, folder, export_result)

        if node == "无前后台管理员SOD问题":
            return self._check_sod(export_result)

        return None

    def _resolve_dev_list(self, folder: Path) -> Optional[Path]:
        """查找开发人员清单（兼容文件名编码损坏）。"""
        dev = self._find_file_by_content(
            folder, "开发人员清单.xlsx",
            required_columns=["姓名"],
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
            if folder.name == "OPM":
                dev_df = pd.read_excel(dev_list, sheet_name="开发人员名单")
                dev_df = dev_df.rename(columns={"姓名（中文或英文全名）": "姓名"})
                dev_df["域账号"] = ""
            else:
                dev_df = pd.read_excel(dev_list, sheet_name="Sheet1")

            audit_df = export_result.df
            violations = []
            for _, d in dev_df.iterrows():
                name = str(d.get("姓名", "")).strip()
                acct = str(d.get("域账号", "")).strip()
                matched = pd.DataFrame()
                if acct:
                    matched = audit_df[audit_df["账号名"].astype(str).str.strip().str.upper() == acct.upper()]
                if name and matched.empty:
                    matched = audit_df[audit_df["账号使用人"].astype(str).str.contains(name, na=False)]
                for _, m in matched.iterrows():
                    if str(m.get("用途", "")).strip() in {"读写", "管理员"}:
                        violations.append(f"{name}({acct}):{m['用途']}")

            if violations:
                return {"status": "未通过", "detail": "; ".join(violations[:5])}
            return {"status": "通过", "detail": "无开发人员持有非只读权限"}
        except Exception as e:
            return {"status": "异常", "detail": str(e)}

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

        app_users = set(df[df["审阅对象"].isin(APP) & df["用途"].isin(ADMIN)]
                        ["账号使用人"].dropna().astype(str).str.strip())
        db_users = set(df[df["审阅对象"].isin(DB) & df["用途"].isin(ADMIN)]
                       ["账号使用人"].dropna().astype(str).str.strip())
        overlaps = (app_users & db_users) - SKIP

        if overlaps:
            return {"status": "未通过",
                    "detail": f"SOD重合: {', '.join(sorted(overlaps))}"}
        return {"status": "通过", "detail": "无前后台管理员SOD重合"}

    # ----------------------------------------------------------
    # 汇总最终结论
    # ----------------------------------------------------------

    @staticmethod
    def _compute_final(
        account_ok: Optional[bool],
        ocr_needed: bool,
        ocr_ok: Optional[bool],
        derived_result: Optional[Dict[str, str]],
    ) -> str:
        # 派生检查优先
        if derived_result:
            s = derived_result.get("status", "")
            if s == "通过":
                return "是"
            else:
                return f"否-{derived_result.get('detail', s)}"

        # 账号不存在 → 失败
        if account_ok is False:
            return "否-账号缺失"

        # 不需 OCR → 账号存在即通过
        if not ocr_needed:
            return "是" if account_ok is True else "待确认"

        # 需要 OCR
        if ocr_ok is True:
            return "是"
        elif ocr_ok is False:
            return "否-截图检查未通过"
        else:
            return "否-截图待检查"

    # ----------------------------------------------------------
    # 构建最终 Excel
    # ----------------------------------------------------------

    def _build_conclusion_df(
        self, wp_df: pd.DataFrame, system_results: Dict[str, Dict[str, Dict]]
    ) -> pd.DataFrame:
        rows = []
        for _, wp_row in wp_df.iterrows():
            system = str(wp_row.get("System", "")).strip()
            row = {c: wp_row.get(c, "") for c in wp_df.columns}
            nc = system_results.get(system, {})

            for node in AuditNodes.ALL:
                row[node] = nc.get(node, {}).get("final", "未评估")

            passed = sum(1 for n in AuditNodes.ALL
                        if str(row.get(n, "")).startswith("是"))
            total = sum(1 for n in AuditNodes.ALL
                       if not str(row.get(n, "")).startswith("N/A"))
            row["通过/需检查"] = f"{passed}/{total}"
            rows.append(row)
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
