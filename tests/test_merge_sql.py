from pathlib import Path
import tempfile
import unittest

import pandas as pd

from sql_healthcheck.merge_sql import merge_sql_csv
from sql_healthcheck.merge_sql import merge_sql_root_healthcheck
from sql_healthcheck.merge_sql import merge_sql_root_csv


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

if __name__ == "__main__":
    unittest.main()
