import unittest
import excelwriter


class TestPptReader(unittest.TestCase):

    def test_Init_WritesColumns_MatchesMap(self):
        writer = excelwriter.ExcelWriter("", {})
        self.assertEqual("Keyword", writer.ws["A1"].value)
        self.assertEqual("Filename", writer.ws["B1"].value)
        self.assertEqual("Count", writer.ws["C1"].value)
        self.assertEqual("Filepath", writer.ws["D1"].value)
        self.assertEqual("Slide Indices", writer.ws["E1"].value)


if __name__ == "__main__":
    unittest.main()