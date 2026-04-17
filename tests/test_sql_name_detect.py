import unittest

from sql_healthcheck.name_detect import extract_sheet_name


class SqlNameDetectTests(unittest.TestCase):
    def test_uses_previous_part_when_filename_ends_with_long_timestamp(self):
        self.assertEqual(extract_sheet_name("DB1 - Database Info - 20260417103000.csv"), "Database Info")

    def test_uses_last_part_when_it_contains_letters(self):
        self.assertEqual(extract_sheet_name("DB1 - Wait Events.csv"), "Wait Events")

    def test_short_numeric_suffix_scans_left_for_letters(self):
        self.assertEqual(extract_sheet_name("DB1 - Sessions - 35.csv"), "Sessions")

    def test_invalid_excel_sheet_chars_follow_reference_logic(self):
        self.assertEqual(extract_sheet_name("DB1 - A:B/C?D*E[1].csv"), "A:B/C?D*E[1]")

    def test_numeric_only_name_falls_back_to_safe_sheet_name(self):
        self.assertEqual(extract_sheet_name("123456.csv"), "123456")


if __name__ == "__main__":
    unittest.main()
