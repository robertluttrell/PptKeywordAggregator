import unittest
import pptreader
import os
import KeywordFilePresence as kfp


class TestPptReader(unittest.TestCase):

    def test_Add_Index_IndexListAppended(self):
        new_kfp = kfp.KeywordFilePresence("")
        new_kfp.add_index(1)
        new_kfp.add_index(2)
        new_kfp.add_index(3)
        self.assertListEqual(new_kfp.index_list, [1, 2, 3])


if __name__ == "__main__":
    unittest.main()
