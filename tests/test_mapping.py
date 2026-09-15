from pathlib import Path
import unittest

from app_logic import DATA_GUARD_PARAMETER_NAMES, _filter_parameter_rows
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

    def test_filters_data_guard_parameters_from_all_parameters(self):
        content = TableContent(
            source_path=Path("all_parameters.html"),
            rows=[
                ["NAME", "VALUE"],
                ["LOG_ARCHIVE_CONFIG", "DG_CONFIG=(FSSPRO,FSSDR)"],
                ["LOG_ARCHIVE_DEST_1", "LOCATION=USE_DB_RECOVERY_FILE_DEST VALID_FOR=(ALL_LOGFILES,ALL_ROLES) DB_UNIQUE_NAME=FSSPRO"],
                ["LOG_ARCHIVE_DEST_2", "SERVICE=FSSDR LGWR ASYNC VALID_FOR=(ONLINE_LOGFILES,PRIMARY_ROLE) DB_UNIQUE_NAME=FSSDR"],
                ["LOG_ARCHIVE_DEST_3", "nan"],
                ["LOG_ARCHIVE_DEST_4", "nan"],
                ["LOG_ARCHIVE_DEST_STATE_1", "enable"],
                ["LOG_ARCHIVE_DEST_STATE_2", "ENABLE"],
                ["LOG_ARCHIVE_DEST_STATE_3", "enable"],
                ["LOG_ARCHIVE_DEST_STATE_4", "enable"],
                ["dg_broker_config_file1", "/u01/app/19.0.0/oracle/dbs/dr1FSSPRO.dat"],
                ["dg_broker_config_file2", "/u01/app/19.0.0/oracle/dbs/dr2FSSPRO.dat"],
                ["fal_client", "FSSPRO"],
                ["fal_server", "FSSDR"],
                ["control_files", "/u01/control01.ctl,/u02/control02.ctl"],
            ],
            logical_key="all_parameters",
            keys={"all_parameters"},
        )

        filtered = _filter_parameter_rows(content, DATA_GUARD_PARAMETER_NAMES, output_headers=["NAME", "VALUE"])

        self.assertEqual(filtered.rows[0], ["NAME", "VALUE"])
        self.assertEqual([row[0] for row in filtered.rows[1:]], [
            "LOG_ARCHIVE_CONFIG",
            "LOG_ARCHIVE_DEST_1",
            "LOG_ARCHIVE_DEST_2",
            "LOG_ARCHIVE_DEST_3",
            "LOG_ARCHIVE_DEST_4",
            "LOG_ARCHIVE_DEST_STATE_1",
            "LOG_ARCHIVE_DEST_STATE_2",
            "LOG_ARCHIVE_DEST_STATE_3",
            "LOG_ARCHIVE_DEST_STATE_4",
            "dg_broker_config_file1",
            "dg_broker_config_file2",
            "fal_client",
            "fal_server",
        ])
        self.assertEqual(filtered.rows[3][1], "SERVICE=FSSDR LGWR ASYNC VALID_FOR=(ONLINE_LOGFILES,PRIMARY_ROLE) DB_UNIQUE_NAME=FSSDR")

    def test_control_files_still_split_values(self):
        content = TableContent(
            source_path=Path("all_parameters.html"),
            rows=[
                ["NAME", "VALUE"],
                ["control_files", "/u01/control01.ctl, /u02/control02.ctl"],
            ],
        )

        filtered = _filter_parameter_rows(content, {"control_files"}, split_values=True)

        self.assertEqual(filtered.rows, [
            ["PARAMETER", "VALUE"],
            ["control_files", "/u01/control01.ctl"],
            ["control_files", "/u02/control02.ctl"],
        ])


if __name__ == "__main__":
    unittest.main()

