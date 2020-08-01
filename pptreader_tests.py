import unittest
import pptreader


class TestPptReader(unittest.TestCase):

    def test_get_ppt_path_list(self):
        path_list = pptreader.PptReader.get_ppt_path_list()


if __name__ == "__main__":
    unittest.main()