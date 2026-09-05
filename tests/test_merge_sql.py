from pathlib import Path
import tempfile
import unittest

import pandas as pd

from sql_healthcheck.merge_sql import merge_sql_csv
from sql_healthcheck.merge_sql import merge_sql_root_healthcheck
from sql_healthcheck.merge_sql import merge_sql_root_csv
from rpwithchart import build_sql_summary_rows


class MergeSqlTests(unittest.TestCase):
    def test_empty_folder_returns_none(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "out.xlsx"
            self.assertIsNone(merge_sql_csv(temp_dir, output))
            self.assertFalse(output.exists())

    def test_duplicate_sheet_names_are_concatenated(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            folder = Path(temp_dir)
            pd.DataFrame({"id": [1], "value": ["a"]}).to_csv(folder / "DB - Sessions - 1.csv", index=False)
            pd.DataFrame({"id": [2], "value": ["b"]}).to_csv(folder / "DB - Sessions - 2.csv", index=False)

            output = folder / "merged.xlsx"
            result = merge_sql_csv(folder, output)

            self.assertEqual(result, str(output))
            workbook = pd.read_excel(output, sheet_name=None)
            self.assertIn("Sessions", workbook)
            self.assertEqual(len(workbook["Sessions"]), 2)

    def test_multiple_sheet_names_are_written(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            folder = Path(temp_dir)
            pd.DataFrame({"id": [1]}).to_csv(folder / "DB - Database Info.csv", index=False)
            pd.DataFrame({"id": [2]}).to_csv(folder / "DB - Wait Events.csv", index=False)

            output = folder / "merged.xlsx"
            merge_sql_csv(folder, output)
            workbook = pd.read_excel(output, sheet_name=None)

            self.assertIn("Database Info", workbook)
            self.assertIn("Wait Events", workbook)

    def test_volume_info_uses_only_latest_snapshot(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            folder = Path(temp_dir)
            old_snapshot = pd.DataFrame({
                "volume_mount_point": ["D:\\"],
                "Total Size (GB)": [199.87],
                "Available Size (GB)": [80.18],
                "Space Free %": [40.12],
            })
            new_snapshot = pd.DataFrame({
                "volume_mount_point": ["D:\\"],
                "Total Size (GB)": [199.87],
                "Available Size (GB)": [77.76],
                "Space Free %": [38.90],
            })
            old_snapshot.to_csv(
                folder / "SERVER$INSTANCE-DQ-26-Volume Info-202606091531523152.csv",
                index=False,
            )
            new_snapshot.to_csv(
                folder / "SERVER$INSTANCE-DQ-26-Volume Info-202608280921302130.csv",
                index=False,
            )

            output = folder / "merged.xlsx"
            merge_sql_csv(folder, output)
            volume_info = pd.read_excel(output, sheet_name="Volume Info")

            self.assertEqual(len(volume_info), 1)
            self.assertEqual(volume_info.loc[0, "volume_mount_point"], "D:\\")
            self.assertAlmostEqual(volume_info.loc[0, "Available Size (GB)"], 77.76)
            self.assertAlmostEqual(volume_info.loc[0, "Space Free %"], 38.90)

    def test_volume_info_keeps_all_latest_files_and_removes_exact_duplicates(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            folder = Path(temp_dir)
            first = pd.DataFrame({
                "volume_mount_point": ["D:\\", "E:\\"],
                "Available Size (GB)": [77.76, 120.0],
            })
            second = pd.DataFrame({
                "volume_mount_point": ["D:\\", "F:\\"],
                "Available Size (GB)": [77.76, 200.0],
            })
            first.to_csv(
                folder / "SERVER$INSTANCE-DQ-26-Volume Info-202608280921302130.csv",
                index=False,
            )
            second.to_csv(
                folder / "SERVER$INSTANCE-DQ-026-Volume Info-202608280921302130.csv",
                index=False,
            )

            output = folder / "merged.xlsx"
            merge_sql_csv(folder, output)
            volume_info = pd.read_excel(output, sheet_name="Volume Info")

            self.assertCountEqual(volume_info["volume_mount_point"].tolist(), ["D:\\", "E:\\", "F:\\"])

    def test_root_merge_combines_db_subfolders(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            db_one = root / "DB_ONE"
            db_two = root / "DB_TWO"
            db_one.mkdir()
            db_two.mkdir()
            pd.DataFrame({"id": [1], "value": ["a"]}).to_csv(db_one / "DB - Sessions - 1.csv", index=False)
            pd.DataFrame({"id": [2], "value": ["b"]}).to_csv(db_two / "DB - Sessions - 2.csv", index=False)

            output = root / "merged_healthcheck_info.xlsx"
            result = merge_sql_root_csv(root, output)

            self.assertEqual(result, str(output))
            workbook = pd.read_excel(output, sheet_name=None)
            self.assertIn("Sessions", workbook)
            self.assertEqual(len(workbook["Sessions"]), 2)

    def test_healthcheck_merge_detects_db_subfolder_csv_shape(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            db_one = root / "DB_ONE"
            db_one.mkdir()
            pd.DataFrame({"id": [1]}).to_csv(db_one / "DB - Database Info.csv", index=False)

            output = root / "merged_healthcheck_info.xlsx"
            result = merge_sql_root_healthcheck(root, output)

            self.assertEqual(result, str(output))
            workbook = pd.read_excel(output, sheet_name=None)
            self.assertIn("Database Info", workbook)

    def test_healthcheck_merge_accepts_selected_db_folder_with_direct_csv(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            db_folder = Path(temp_dir)
            pd.DataFrame({"id": [1]}).to_csv(db_folder / "DB - Database Info.csv", index=False)

            output = db_folder / "merged_healthcheck_info.xlsx"
            result = merge_sql_root_healthcheck(db_folder, output)

            self.assertEqual(result, str(output))
            workbook = pd.read_excel(output, sheet_name=None)
            self.assertIn("Database Info", workbook)

    def test_builds_sql_summary_rows_in_english(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "merged.xlsx"
            with pd.ExcelWriter(output) as writer:
                pd.DataFrame(
                    {
                        "volume_mount_point": ["M:\\"],
                        "Available Size (GB)": [164.83],
                        "Space Free %": [9.5],
                    }
                ).to_excel(writer, sheet_name="Volume Info", index=False)
                pd.DataFrame({"Database Name": ["master"], "Short Query Text": ["SELECT 1"]}).to_excel(
                    writer,
                    sheet_name="Top Worker Time Queries",
                    index=False,
                )
                pd.DataFrame({"Database Name": ["app"], "Table Name": ["dbo.Customer"]}).to_excel(
                    writer,
                    sheet_name="Missing Indexes",
                    index=False,
                )

            with pd.ExcelFile(output) as workbook:
                rows = build_sql_summary_rows(workbook, language="en")

        self.assertEqual(rows[0]["category"], "3.1 Hard disk capacity")
        self.assertIn("Server hard disk", rows[0]["assessment"])
        self.assertEqual(rows[0]["risk"], "May affect storage capacity and operations if disk usage continues to grow.")
        self.assertEqual(rows[1]["category"], "5.1 Queries consuming the most resources")
        self.assertEqual(rows[2]["recommendation"], "Create additional indexes for the tables suggested in section 5.2.")

if __name__ == "__main__":
    unittest.main()
