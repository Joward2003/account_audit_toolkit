import unittest
from pathlib import Path
from types import SimpleNamespace

import pandas as pd

from audit_nodes import AuditNodes
from field_translator import FieldTranslator
from main import AuditConclusionGenerator
from rule_config import load_system_rule_overrides
from sid_utils import normalize_sid
from system_folder_mapper import build_manual_mapping, load_mapping


class SidIdentityTests(unittest.TestCase):
    def test_numeric_excel_sid_is_canonicalized(self):
        self.assertEqual(normalize_sid(1263.0), "1263")
        self.assertEqual(normalize_sid("  ab-01 "), "AB-01")

    def test_folder_resolution_is_exact_not_prefix_based(self):
        folder_map = {"123": "/supporting/123"}
        self.assertEqual(
            AuditConclusionGenerator._resolve_folder("12", folder_map),
            "",
        )
        self.assertEqual(
            AuditConclusionGenerator._resolve_folder("123", folder_map),
            "/supporting/123",
        )

    def test_same_system_display_name_with_different_sids_does_not_collide(self):
        wp_df = pd.DataFrame([
            {"System": "同名系统", "SID": "SID-A", "Category": "SOX System"},
            {"System": "同名系统", "SID": "SID-B", "Category": "SOX System"},
        ])
        results = {
            "SID-A": {
                AuditNodes.APP_ADMIN: {
                    "scope": "是",
                    "final": "是-A",
                    "final_status": "通过",
                },
            },
            "SID-B": {
                AuditNodes.APP_ADMIN: {
                    "scope": "是",
                    "final": "否-B",
                    "final_status": "未通过",
                },
            },
        }

        output = AuditConclusionGenerator()._build_test_result_df(wp_df, results)
        node_rows = output[output["检查节点"] == AuditNodes.APP_ADMIN]
        by_sid = dict(zip(node_rows["SID"], node_rows["结论"]))
        self.assertEqual(by_sid, {"SID-A": "是-A", "SID-B": "否-B"})

    def test_duplicate_sid_is_rejected_before_results_can_overwrite(self):
        scope_results = [
            SimpleNamespace(facts=SimpleNamespace(sid="1001", system="系统A")),
            SimpleNamespace(facts=SimpleNamespace(sid=1001.0, system="系统B")),
        ]
        with self.assertRaisesRegex(ValueError, "SID 重复"):
            AuditConclusionGenerator._validate_scope_sid_identity(scope_results)

    def test_special_system_flags_are_read_by_sid_not_display_name(self):
        translator = FieldTranslator()
        translator.system_rules = {
            "SID-FLAG": {
                "is_compliance_support_platform": "是",
                "is_pms_vms": "是",
            }
        }
        flagged = translator.translate_row(pd.Series({
            "Category": "Utility Tool",
            "System": "任意展示名称",
            "SID": "sid-flag",
        }))
        same_name_without_sid_flag = translator.translate_row(pd.Series({
            "Category": "Utility Tool",
            "System": "PMS/VMS",
            "SID": "SID-OTHER",
        }))

        self.assertTrue(flagged.is_compliance_support_platform)
        self.assertTrue(flagged.is_pms_vms)
        self.assertFalse(same_name_without_sid_flag.is_compliance_support_platform)
        self.assertFalse(same_name_without_sid_flag.is_pms_vms)

    def test_generated_sid_configs_have_no_duplicate_keys_and_mark_blank_sid_rows(self):
        config_dir = Path(__file__).resolve().parents[1] / "configs"
        unique_sid_files = [
            "系统文件夹映射表.xlsx",
            "系统规则特殊覆盖表.xlsx",
            "人工覆盖表（预留）.xlsx",
        ]
        for filename in unique_sid_files:
            df = pd.read_excel(config_dir / filename, dtype=str, keep_default_na=False)
            keys = [normalize_sid(value) for value in df["SID"]]
            nonblank_keys = [key for key in keys if key]
            self.assertEqual(len(nonblank_keys), len(set(nonblank_keys)), filename)
            blank_rows = df[[not key for key in keys]]
            if not blank_rows.empty:
                self.assertIn("备注", blank_rows.columns, filename)
                self.assertTrue(
                    blank_rows["备注"].astype(str).str.contains("暂无SID").all(),
                    filename,
                )

        roles = pd.read_excel(
            config_dir / "应用层管理员角色配置表.xlsx",
            dtype=str,
            keep_default_na=False,
        )
        role_keys = [
            (normalize_sid(row["SID"]), str(row["管理员角色关键词"]).strip().upper())
            for _, row in roles.iterrows()
        ]
        self.assertTrue(all(sid and role for sid, role in role_keys))
        self.assertEqual(len(role_keys), len(set(role_keys)))

    def test_blank_sid_placeholders_are_not_loaded_as_runtime_rules(self):
        config_dir = Path(__file__).resolve().parents[1] / "configs"
        folder_mapping = build_manual_mapping(
            load_mapping(config_dir / "系统文件夹映射表.xlsx")
        )
        self.assertEqual(len(folder_mapping), 41)
        self.assertNotIn("", folder_mapping)

        load_system_rule_overrides.cache_clear()
        rules = load_system_rule_overrides()
        self.assertEqual(len(rules), 6)
        self.assertNotIn("", rules)


if __name__ == "__main__":
    unittest.main()
