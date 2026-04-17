import unittest

from utils.normalize import normalize_key, strip_chart_suffix


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


if __name__ == "__main__":
    unittest.main()

