import unittest

from utils.normalize import content_key_aliases, normalize_key, strip_chart_suffix


class NormalizeTests(unittest.TestCase):
    def test_normalizes_titles_and_placeholders(self):
        self.assertEqual(
            normalize_key("Tables without primary key constraints"),
            "tables_without_primary_key_constraints",
        )
        self.assertEqual(normalize_key("<tbs_usage>"), "tbs_usage")

    def test_strips_chart_suffixes(self):
        self.assertEqual(
            strip_chart_suffix("buffer_cache_hit_ratio_line_chart"),
            "buffer_cache_hit_ratio",
        )

    def test_aliases_history_day_count_after_between_suffix(self):
        aliases = content_key_aliases(
            "ash_top_sql_for_instance_1_for_9_days_of_history_between_2026_06_07t15_08_04_and_2026_06_16t15_08_04"
        )

        self.assertIn("ash_top_sql_for_instance_1_for_days_of_history", aliases)
        self.assertIn("ash_top_sql_for_instance", aliases)


if __name__ == "__main__":
    unittest.main()

