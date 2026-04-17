import unittest

from bs4 import BeautifulSoup

from extraction.table_extractor import html_table_to_matrix


class TableExtractorTests(unittest.TestCase):
    def test_table_cells_are_cleaned_and_padded(self):
        soup = BeautifulSoup(
            """
            <table>
              <tr><th>A</th><th>B</th><th>C</th></tr>
              <tr><td> one </td><td>two</td></tr>
            </table>
            """,
            "html.parser",
        )

        rows = html_table_to_matrix(soup.find("table"))

        self.assertEqual(rows, [["A", "B", "C"], ["one", "two", ""]])


if __name__ == "__main__":
    unittest.main()

