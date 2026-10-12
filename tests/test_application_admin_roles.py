import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from application_admin_role_config import match_role_keyword
from evidence_checker import EvidenceChecker, ImageBatch, OCRPage, TextBox
from main import AuditConclusionGenerator


class ApplicationAdminRoleTests(unittest.TestCase):
    def test_fuzzy_keyword_matches_role_with_prefix(self):
        matched, method = match_role_keyword("CSSC", "用户管理-CSSC", "fuzzy")
        self.assertTrue(matched)
        self.assertEqual(method, "模糊包含匹配")

    def test_exact_keyword_does_not_match_role_with_prefix(self):
        matched, _ = match_role_keyword("CSSC", "用户管理-CSSC", "exact")
        self.assertFalse(matched)

    def test_extracts_values_below_role_header_only(self):
        page = OCRPage(
            text_blocks=[
                TextBox("用户", (20, 0, 50, 10), 0.99),
                TextBox("角色", (200, 0, 240, 10), 0.99),
                TextBox("alice", (20, 30, 60, 40), 0.99),
                TextBox("用户管理-CSSC", (200, 30, 310, 40), 0.99),
            ],
            full_text="用户\n角色\nalice\n用户管理-CSSC",
            image_index=0,
        )
        values = EvidenceChecker().extract_column_values(
            [page], "应用层管理员角色完整性", "role_column"
        )
        self.assertEqual(values, ["用户管理-CSSC"])

    def test_role_completeness_uses_sid_rule_and_role_column_ocr(self):
        page = OCRPage(
            text_blocks=[
                TextBox("角色", (200, 0, 240, 10), 0.99),
                TextBox("用户管理-CSSC", (200, 30, 310, 40), 0.99),
            ],
            full_text="角色\n用户管理-CSSC",
            image_index=0,
        )
        generator = AuditConclusionGenerator()
        generator.evidence_loader = Mock(return_value=None)
        generator.evidence_loader.load_for_node.return_value = ImageBatch(
            node="应用层管理员角色完整性",
            sheet_name="应用层管理员角色截图",
            images=[], image_count=1, source_file="supporting.xlsx",
        )
        generator.ocr_extractor = Mock()
        generator.ocr_extractor.extract.return_value = [page]

        with patch("main.load_application_admin_role_rules", return_value={
            "SID001": [{
                "sid": "SID001", "role_keyword": "CSSC", "match_mode": "fuzzy", "note": "",
            }],
        }):
            result = generator._check_application_admin_role_completeness(
                "SID001", Path("supporting.xlsx")
            )

        self.assertEqual(result["status"], "通过")
        self.assertEqual(result["records"][0]["actual_role"], "用户管理-CSSC")


if __name__ == "__main__":
    unittest.main()
