import unittest
import pptreader
import os


class TestPptReader(unittest.TestCase):

    def test_GetPptPathList_EmptyDir_EmptyArray(self):
        path_list = pptreader.PptReader.get_ppt_path_list(os.getcwd() + os.path.sep + "testdir2")
        self.assertListEqual(path_list, [])

    def test_GetPptPathList_PptsInDir_ExactArray(self):
        path_list = pptreader.PptReader.get_ppt_path_list(os.getcwd() + os.path.sep + "testdir1")
        self.assertEqual(len(path_list), 3)
        self.assertListEqual(path_list, [os.getcwd() + os.path.sep + "testdir1" + os.path.sep + "Presentation1.pptx",
                                         os.getcwd() + os.path.sep + "testdir1" + os.path.sep + "Presentation2.pptx",
                                         os.getcwd() + os.path.sep + "testdir1" + os.path.sep + "Presentation3.pptx"])

    def test_GetPptPathList_EmptyDir_WorkingDirUnmodified(self):
        orig_path = os.getcwd()
        pptreader.PptReader.get_ppt_path_list(os.getcwd() + os.path.sep + "testdir1")
        self.assertEqual(os.getcwd(), orig_path)

    def test_GetPptPathList_NonexistentDir_RaisesException(self):
        with self.assertRaises(FileNotFoundError):
            pptreader.PptReader.get_ppt_path_list(os.getcwd() + "NonexistentDirPath")


if __name__ == "__main__":
    unittest.main()