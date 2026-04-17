from pathlib import Path
import unittest

from mapping.content_registry import ContentRegistry
from mapping.mapper import resolve_mappings
from models import GenerationReport, MappingRule, TableContent


class MappingTests(unittest.TestCase):
    def test_resolves_table_by_normalized_source_key(self):
        content = TableContent(
            source_path=Path("tablespace.html"),
            rows=[["A"], ["B"]],
            logical_key="tablespace_usage",
            keys={"Tablespace Usage", "tablespace_usage"},
        )
        registry = ContentRegistry([content])
        report = GenerationReport()
        rule = MappingRule(
            placeholder="<tbs_usage>",
            source_key="Tablespace Usage",
            content_type="table",
        )

        resolved = resolve_mappings([rule], registry, report)

        self.assertIn("<tbs_usage>", resolved)
        self.assertEqual(report.missing_content, [])


if __name__ == "__main__":
    unittest.main()

