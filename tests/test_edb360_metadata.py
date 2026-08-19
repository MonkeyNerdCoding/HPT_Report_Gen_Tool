from pathlib import Path
import unittest

from extraction.edb360_metadata import build_edb360_text_mapping, extract_edb360_metadata
from extraction.edb360_assessment import build_edb360_assessment_mapping


SAMPLE_EDB360_ROOT = Path(r"D:\HPT\CGV\edb360_082026")


@unittest.skipUnless(SAMPLE_EDB360_ROOT.is_dir(), "local EDB360 sample is not available")
class Edb360MetadataTests(unittest.TestCase):
    def test_extracts_cover_metadata_from_edb360(self):
        metadata = extract_edb360_metadata(SAMPLE_EDB360_ROOT)

        self.assertEqual(metadata["database_name"], "GCGVDB")
        self.assertEqual(metadata["oracle_version"], "12.1.0.2.0")
        self.assertEqual(metadata["collection_date"], "Aug-2026")
        self.assertIn("gcgvdb", metadata["instance_summary"])

    def test_builds_defaults_for_manual_signature_fields(self):
        mapping = build_edb360_text_mapping({"database_name": "GCGVDB"})

        self.assertEqual(mapping["{{database_name}}"], "GCGVDB")
        self.assertIn("{{creator}}", mapping)
        self.assertIn("{{approver}}", mapping)

    def test_builds_dynamic_assessment_text_from_edb360(self):
        mapping = build_edb360_assessment_mapping(SAMPLE_EDB360_ROOT)

        self.assertIn("4 control files", mapping["{{assessment_control_file}}"])
        self.assertIn("redo log group đang được multiplexing", mapping["{{assessment_redo_log}}"])
        self.assertIn("SGA 180.000G/instance", mapping["{{assessment_memory_configuration}}"])
        self.assertIn("1146 bảng không có index", mapping["{{assessment_no_index}}"])
        self.assertIn("1678 bảng không có khóa chính", mapping["{{assessment_no_pk}}"])


if __name__ == "__main__":
    unittest.main()
