from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd

from audit_nodes import AuditNodes
from config import AuditConfig
from decisions import Decision
from evidence_task_router import EvidenceTaskRouter
from field_translator import FieldTranslator
from pipeline import AccountAuditPipeline
from rules_scope import AccountAuditScopeEngine, ScopeResult

def _now_iso() -> str:
    """生成稳定、可读的时间戳。"""
    return datetime.now().replace(microsecond=0).isoformat()


def _make_task_id(system: str, node: str, seq: int) -> str:
    """给后续 workflow / Dify / LangGraph 节点传递用的任务 ID。"""
    base = f"{system}-{node}-{seq}"
    base = re.sub(r"\s+", "_", base)
    base = re.sub(r"[^\w\u4e00-\u9fff.-]+", "_", base)
    return base.strip("_")


class AccountAuditWorkflowExporter:
    """
    将 AccountAuditScopeEngine 的判断结果，拆成三类输出：

    1. detail_df：
       给审计/开发人员看的完整明细，保留原因、warning 和中间事实。

    2. program_template_df：
       给底稿/模板使用的简洁宽表。
       只保留 System + 各检查节点的最终 label，不包含原因列。

    3. todo_df / payload_json：
       给后续 OCR、文件夹映射、证据检查节点使用的任务清单。
       一行一个待执行任务，不再携带原因列，避免后续脚本反复解析底稿。
    """

    def __init__(self, config: Optional[AuditConfig] = None):
        self.config = config or AuditConfig()
        self.pipeline = AccountAuditPipeline(self.config)
        self.translator = FieldTranslator(self.config)
        self.scope_engine = AccountAuditScopeEngine()
        self.task_router = EvidenceTaskRouter()

    def build_scope_results(self, df: pd.DataFrame) -> List[ScopeResult]:
        results: List[ScopeResult] = []
        for i, (_, row) in enumerate(df.iterrows(), start=2):
            facts = self.translator.translate_row(row, row_number=i)
            results.append(self.scope_engine.determine(facts))
        return results

    def build_detail_df(self, df: pd.DataFrame) -> pd.DataFrame:
        """保留原始完整输出：中间事实 + label + 是否需凭证 + ReferTo + 原因 + Warnings。"""
        return self.pipeline.run(df)

    def build_program_template_df(self, scope_results: List[ScopeResult]) -> pd.DataFrame:
        """
        生成不含原因的程序模板宽表。

        特点：
        - 一行一个 System。
        - 只保留审计节点最终 label。
        - 不保留“原因”列。
        - 可以直接作为人工复核 / 底稿粘贴的中间表。
        """
        rows: List[Dict[str, Any]] = []

        for result in scope_results:
            facts = result.facts
            row: Dict[str, Any] = {
                "Excel行号": facts.row_number,
                "Category": facts.category,
                "System": facts.system,
                "服务器类型": facts.server_type,
            }

            for node in AuditNodes.ALL:
                decision = result.decisions.get(node, Decision.unknown("节点无判断结果。"))
                row[node] = decision.label

            row["Warnings"] = " | ".join(result.warnings)
            rows.append(row)

        return pd.DataFrame(rows)

    def build_todo_df(self, scope_results: List[ScopeResult]) -> pd.DataFrame:
        """
        生成一行一个任务的 to-do list。

        只放后续执行需要的列，不放原因。
        后续节点可以基于：
        - System 映射到系统文件夹；
        - Node / EvidenceType 映射到标准文件命名；
        - ReferTo 判断是否要跳到公共系统凭证；
        - ExpectedKeywords / ExpectedFields 辅助 OCR 与 set 匹配。
        """
        rows: List[Dict[str, Any]] = []

        for result in scope_results:
            facts = result.facts
            tasks = self.task_router.build_tasks(result)

            for seq, task in enumerate(tasks, start=1):
                target_system = task.refer_to or task.system
                rows.append(
                    {
                        "TaskID": _make_task_id(facts.system, task.node, seq),
                        "Excel行号": facts.row_number,
                        "Category": facts.category,
                        "System": facts.system,
                        "TargetSystem": target_system,
                        "检查节点": task.node,
                        "Scope结果": task.expected_label,
                        "EvidenceType": task.evidence_type,
                        "ReferTo": task.refer_to or "",
                        "ExpectedKeywords": " | ".join(task.expected_keywords),
                        "ExpectedFields": " | ".join(task.expected_fields),
                        "Status": "TODO",
                        "服务器类型": facts.server_type,
                    }
                )

        return pd.DataFrame(rows)

    def build_payload(self, scope_results: List[ScopeResult], source_name: str = "") -> Dict[str, Any]:
        """
        生成标准 JSON payload。

        注意：
        - 这个是“数据实例”，用于其他脚本/节点直接消费。
        - 对应的结构约束由 build_json_schema() 输出。
        """
        systems: List[Dict[str, Any]] = []

        for result in scope_results:
            facts = result.facts
            tasks_payload: List[Dict[str, Any]] = []

            tasks = self.task_router.build_tasks(result)
            for seq, task in enumerate(tasks, start=1):
                target_system = task.refer_to or task.system
                tasks_payload.append(
                    {
                        "task_id": _make_task_id(facts.system, task.node, seq),
                        "system": task.system,
                        "target_system": target_system,
                        "server_type": facts.server_type,
                        "node": task.node,
                        "scope_label": task.expected_label,
                        "evidence_type": task.evidence_type,
                        "refer_to": task.refer_to,
                        "expected_keywords": task.expected_keywords,
                        "expected_fields": task.expected_fields,
                        "status": "TODO",
                    }
                )

            scope_payload = {
                node: {
                    "label": result.decisions.get(node, Decision.unknown()).label,
                    "requires_evidence": result.decisions.get(node, Decision.unknown()).requires_evidence,
                    "refer_to": result.decisions.get(node, Decision.unknown()).refer_to,
                }
                for node in AuditNodes.ALL
            }

            systems.append(
                {
                    "excel_row": facts.row_number,
                    "category": facts.category,
                    "system": facts.system,
                    "folder_key": facts.system,
                    "facts": {
                        "server_type": facts.server_type,
                        "uses_full_dba_paas": facts.uses_full_dba_paas,
                        "is_app_server_containerized": facts.is_app_server_containerized,
                        "no_app_layer": facts.no_app_layer,
                        "no_app_layer_user": facts.no_app_layer_user,
                        "no_backend": facts.no_backend,
                        "no_database_or_db_server": facts.no_database_or_db_server,
                        "is_sox_system": facts.is_sox_system,
                        "is_compliance_support_platform": facts.is_compliance_support_platform,
                        "is_utility_tool": facts.is_utility_tool,
                        "is_other": facts.is_other,
                        "is_pms_vms": facts.is_pms_vms,
                    },
                    "scope": scope_payload,
                    "tasks": tasks_payload,
                    "warnings": result.warnings,
                }
            )

        return {
            "schema_version": "account_audit_todo.v1",
            "generated_at": _now_iso(),
            "source_name": source_name,
            "systems": systems,
        }

    @staticmethod
    def build_json_schema() -> Dict[str, Any]:
        """
        生成 JSON Schema。

        这个文件给其他脚本、Dify 工作流节点、LangGraph 节点做输入校验。
        """
        return {
            "$schema": "https://json-schema.org/draft/2020-12/schema",
            "$id": "https://local.audit/schema/account_audit_todo.v1.json",
            "title": "Account Audit To-do Payload",
            "type": "object",
            "required": ["schema_version", "generated_at", "systems"],
            "properties": {
                "schema_version": {
                    "type": "string",
                    "const": "account_audit_todo.v1",
                },
                "generated_at": {
                    "type": "string",
                    "description": "ISO format datetime generated by exporter.",
                },
                "source_name": {
                    "type": "string",
                    "description": "Original Excel/CSV file name or business batch name.",
                },
                "systems": {
                    "type": "array",
                    "items": {"$ref": "#/$defs/system_item"},
                },
            },
            "$defs": {
                "system_item": {
                    "type": "object",
                    "required": ["excel_row", "category", "system", "folder_key", "scope", "tasks"],
                    "properties": {
                        "excel_row": {"type": ["integer", "null"]},
                        "category": {"type": "string"},
                        "system": {"type": "string"},
                        "folder_key": {
                            "type": "string",
                            "description": "Default equals system. Later scripts may map this to actual folder name.",
                        },
                        "facts": {"type": "object"},
                        "scope": {
                            "type": "object",
                            "additionalProperties": {"$ref": "#/$defs/scope_decision"},
                        },
                        "tasks": {
                            "type": "array",
                            "items": {"$ref": "#/$defs/task_item"},
                        },
                        "warnings": {
                            "type": "array",
                            "items": {"type": "string"},
                        },
                    },
                    "additionalProperties": False,
                },
                "scope_decision": {
                    "type": "object",
                    "required": ["label", "requires_evidence", "refer_to"],
                    "properties": {
                        "label": {
                            "type": "string",
                            "description": "是 / N/A / Refer to XXX / N/A-Refer to XXX / 待确认 / custom label",
                        },
                        "requires_evidence": {"type": "boolean"},
                        "refer_to": {"type": ["string", "null"]},
                    },
                    "additionalProperties": False,
                },
                "task_item": {
                    "type": "object",
                    "required": [
                        "task_id",
                        "system",
                        "target_system",
                        "server_type",
                        "node",
                        "scope_label",
                        "evidence_type",
                        "expected_keywords",
                        "expected_fields",
                        "status",
                    ],
                    "properties": {
                        "task_id": {"type": "string"},
                        "system": {"type": "string"},
                        "target_system": {
                            "type": "string",
                            "description": "If refer_to exists, target_system equals refer_to; otherwise equals system.",
                        },
                        "server_type": {
                            "type": "string",
                            "description": "Server type copied from WP, e.g. Linux / Windows / Linux&Windows. Used by execution plan builder.",
                        },
                        "node": {"type": "string"},
                        "scope_label": {"type": "string"},
                        "evidence_type": {"type": "string"},
                        "refer_to": {"type": ["string", "null"]},
                        "expected_keywords": {
                            "type": "array",
                            "items": {"type": "string"},
                        },
                        "expected_fields": {
                            "type": "array",
                            "items": {"type": "string"},
                        },
                        "status": {
                            "type": "string",
                            "enum": ["TODO", "RUNNING", "PASS", "FAIL", "REVIEW", "SKIPPED"],
                        },
                    },
                    "additionalProperties": False,
                },
            },
            "additionalProperties": False,
        }

    def export(
        self,
        df: pd.DataFrame,
        output_dir: str | Path = ".",
        base_name: str = "account_audit_todo",
        source_name: str = "",
        include_detail_sheet: bool = True,
    ) -> Dict[str, str]:
        """
        一键导出三类文件：

        - {base_name}.xlsx
          sheet1: todo_list，一行一个待执行任务
          sheet2: program_template，不含原因的宽表模板
          sheet3: detail_with_reasons，可选，保留完整原因和 warning

        - {base_name}.json
          标准任务 payload，给后续脚本/节点消费

        - {base_name}.schema.json
          JSON Schema，给其他节点做结构校验
        """
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)

        scope_results = self.build_scope_results(df)
        todo_df = self.build_todo_df(scope_results)
        template_df = self.build_program_template_df(scope_results)
        detail_df = self.build_detail_df(df) if include_detail_sheet else None

        excel_path = output_path / f"{base_name}.xlsx"
        payload_path = output_path / f"{base_name}.json"
        schema_path = output_path / f"{base_name}.schema.json"

        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            todo_df.to_excel(writer, sheet_name="todo_list", index=False)
            template_df.to_excel(writer, sheet_name="program_template", index=False)
            if include_detail_sheet and detail_df is not None:
                detail_df.to_excel(writer, sheet_name="detail_with_reasons", index=False)

        payload = self.build_payload(scope_results, source_name=source_name)
        payload_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

        schema = self.build_json_schema()
        schema_path.write_text(json.dumps(schema, ensure_ascii=False, indent=2), encoding="utf-8")

        return {
            "excel": str(excel_path),
            "json": str(payload_path),
            "schema": str(schema_path),
        }
