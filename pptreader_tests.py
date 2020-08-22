import unittest
import pptreader
import os


class TestPptReader(unittest.TestCase):

    def test_GetPptPathList_EmptyDir_EmptyArray(self):
        path_list = pptreader.PptReader.get_ppt_path_list(os.getcwd() + os.path.sep + "Test" + os.path.sep + "testdir2")
        self.assertListEqual(path_list, [])

    def test_GetPptPathList_PptsInDir_ExactArray(self):
        path_list = pptreader.PptReader.get_ppt_path_list(os.getcwd() + os.path.sep + "Test" + os.path.sep + "testdir1")
        self.assertEqual(len(path_list), 3)
        self.assertListEqual(path_list,
                             [os.getcwd() + os.path.sep + "Test" + os.path.sep + "testdir1" + os.path.sep + "Presentation1.pptx",
                              os.getcwd() + os.path.sep + "Test" + os.path.sep + "testdir1" + os.path.sep + "Presentation2.pptx",
                              os.getcwd() + os.path.sep + "Test" + os.path.sep + "testdir1" + os.path.sep + "Presentation3.pptx"])

    def test_GetPptPathList_EmptyDir_WorkingDirUnmodified(self):
        orig_path = os.getcwd()
        pptreader.PptReader.get_ppt_path_list(os.getcwd() + os.path.sep + "Test" + os.path.sep + "testdir1")
        self.assertEqual(os.getcwd(), orig_path)

    def test_GetPptPathList_NonexistentDir_RaisesException(self):
        with self.assertRaises(FileNotFoundError):
            pptreader.PptReader.get_ppt_path_list(os.getcwd() + "NonexistentDirPath")

    def test_PptFilesToDict_EmptyList_CreatesEmptyDict(self):
        path_list = []
        reader = pptreader.PptReader("")
        reader.ppt_files_to_dict(path_list)
        self.assertEqual({}, reader.word_dict)

    def test_PptFilesToDict_NoKeywords_CreatesEmptyDict(self):
        test_dir_path = os.getcwd() + os.path.sep + "Test" + os.path.sep + "testdir1"
        path_list = self.get_ppt_file_list(test_dir_path)
        reader = pptreader.PptReader("")

        reader.ppt_files_to_dict(path_list)

        self.assertTrue(len(path_list) > 0)
        self.assertEqual({}, reader.word_dict)

    def test_PptFilesToDict_MultipleKeywords_ReturnsCorrectCount(self):
        test_dir_path = os.getcwd() + os.path.sep + "Test" + os.path.sep + "testdir3"
        path_list = self.get_ppt_file_list(test_dir_path)
        reader = pptreader.PptReader("")

        reader.ppt_files_to_dict(path_list)

        self.assertEqual(4, len(reader.word_dict))
        print(reader.word_dict.keys())
        self.assertTrue("Keyword3" in reader.word_dict.keys())
        self.assertTrue("Keyword4 Keyword5" in reader.word_dict.keys())

    @staticmethod
    def get_ppt_file_list(dir_path):
        path_list = []
        for filename in os.listdir(dir_path):
            if filename.endswith(".pptx"):
                path_list.append(os.path.join(dir_path, filename))

        return path_list


if __name__ == "__main__":
    unittest.main()