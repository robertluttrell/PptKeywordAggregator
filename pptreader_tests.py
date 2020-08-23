import unittest
import pptreader
import os
import keywordfilepresence as kfp


class TestPptReader(unittest.TestCase):

    def test_GetPptPathList_EmptyDir_EmptyArray(self):
        path_list = pptreader.PptReader.get_ppt_path_list(os.getcwd() + os.path.sep + "Test" + os.path.sep + "EmptyDir")
        self.assertListEqual(path_list, [])

    def test_GetPptPathList_PptsInDir_ExactArray(self):
        path_list = pptreader.PptReader.get_ppt_path_list(os.getcwd() + os.path.sep + "Test" + os.path.sep
                                                          + "BlankPptTestFiles")
        self.assertEqual(len(path_list), 3)
        self.assertListEqual(path_list,
                             [os.getcwd() + os.path.sep + "Test" + os.path.sep + "BlankPptTestFiles" + os.path.sep
                              + "OneSlideNoKeywords1.pptx",
                              os.getcwd() + os.path.sep + "Test" + os.path.sep + "BlankPptTestFiles" + os.path.sep
                              + "OneSlideNoKeywords2.pptx",
                              os.getcwd() + os.path.sep + "Test" + os.path.sep + "BlankPptTestFiles" + os.path.sep
                              + "OneSlideNoKeywords3.pptx"])

    def test_GetPptPathList_EmptyDir_WorkingDirUnmodified(self):
        orig_path = os.getcwd()
        pptreader.PptReader.get_ppt_path_list(os.getcwd() + os.path.sep + "Test" + os.path.sep + "BlankPptTestFiles")
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
        test_dir_path = os.getcwd() + os.path.sep + "Test" + os.path.sep + "BlankPptTestFiles"
        path_list = self.get_ppt_file_list(test_dir_path)
        reader = pptreader.PptReader("")

        reader.ppt_files_to_dict(path_list)

        self.assertTrue(len(path_list) > 0)
        self.assertEqual({}, reader.word_dict)

    def test_PptFilesToDict_MultipleKeywords_ReturnsCorrectCount(self):
        test_dir_path = os.getcwd() + os.path.sep + "Test" + os.path.sep + "RepeatedKeywordsDifferentPresentations"
        path_list = self.get_ppt_file_list(test_dir_path)
        reader = pptreader.PptReader("")
        expected_dict = {"Keyword1": [kfp.KeywordFilePresence(path_list[0])],
                         "Keyword2": [kfp.KeywordFilePresence(path_list[0])],
                         "Keyword3": [kfp.KeywordFilePresence(path_list[0]), kfp.KeywordFilePresence(path_list[1])],
                         "Keyword4 Keyword5": [kfp.KeywordFilePresence(path_list[0])]}
        expected_dict["Keyword1"][0].add_index(1)
        expected_dict["Keyword2"][0].add_index(1)
        expected_dict["Keyword3"][0].add_index(1)
        expected_dict["Keyword3"][0].add_index(2)
        expected_dict["Keyword3"][1].add_index(1)
        expected_dict["Keyword3"][1].add_index(2)
        expected_dict["Keyword4 Keyword5"][0].add_index(1)

        reader.ppt_files_to_dict(path_list)

        self.assertEqual(expected_dict, reader.word_dict)

    @staticmethod
    def get_ppt_file_list(dir_path):
        path_list = []
        for filename in os.listdir(dir_path):
            if filename.endswith(".pptx"):
                path_list.append(os.path.join(dir_path, filename))

        return path_list


if __name__ == "__main__":
    unittest.main()