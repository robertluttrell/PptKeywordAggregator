import os
from pptx import Presentation
from KeywordFilePresence import KeywordFilePresence

class PptReader:

    def __init__(self, ppt_dir_path):
        self.ppt_dir_path = ppt_dir_path
        self.word_dict = None

    @staticmethod
    def get_ppt_path_list(ppt_dir_path):
        """
        Creates a list of paths to all the powerpoint files in self.ppt_dir_path
        :return: list of paths
        """
        ppt_path_list = []
        curDir = os.getcwd()

        os.chdir(ppt_dir_path)
        for file in os.listdir(ppt_dir_path):
            if file.endswith(".pptx"):
                ppt_path_list.append(os.path.join(ppt_dir_path, file))

        os.chdir(curDir)
        return ppt_path_list

    def keyword_file_presence_exists(self, keyword, file_path):
        """
        Returns true if the KeywordFilePresence exists, indicating there's already an entry for this keyword and slide
        :param keyword: String representing the keyword
        :param file_path: String representing the path to the PowerPoint file the keyword was found in
        :return: True if exists, else false
        """
        keyword_kfp_list = self.word_dict[keyword]
        for kfp in keyword_kfp_list:
            if kfp.get_file_path() == file_path:
                return True

        return False

    def process_keyword(self, keyword, index, file_path):
        """
        Adds the KeywordFilePresence to self.word_dict
        :param keyword: String representing the keyword
        :param index: Index of slide in powerpoint presentation starting from 1
        :param file_path: String representing path to the PowerPoint presentation containing the slide
        :return:
        """
        if keyword not in self.word_dict:
            self.word_dict[keyword] = []

        word_dict_entry = self.word_dict[keyword]

        if not self.keyword_file_presence_exists(keyword, file_path):
            word_dict_entry.append(KeywordFilePresence(file_path))

        word_dict_entry[-1].add_index(index)

    def process_slide(self, slide, index, file_path):
        """
        Reads a slide and stores keyword data in self.word_dict if a populated notes slide exists
        :param slide: Slide object to be processed
        :param index: Index of slide in powerpoint presentation starting from 1
        :param file_path: String representing path to the PowerPoint presentation containing the slide
        """
        if not slide.has_notes_slide:
            return

        slide_keyword_list = slide.notes_slide.notes_text_frame.text.split(',').strip(' ')
        for keyword in slide_keyword_list:
            self.process_keyword(keyword, index, file_path)


    def ppt_file_to_dict(self, file_path):
        """
        Reads the file and stores keyword data in self.word_dict
        :param file_path: String path to the file to be read
        """
        try:
            file = open(r"file_path", "rb")

        except IOError:
            return

        pres = Presentation(file)
        file.close()

        for i in range(len(pres.slides)):
            self.process_slide(pres.slides[i], i, file_path)

    def ppt_files_to_dict(self, ppt_path_list):
        """
        For each file specified in ppt_path_list, reads the file and stores keyword data in self.word_dict
        :param ppt_path_list: List of strings specifying the powerpoint files to be processed
        :return: df: DataFrame containing keyword data and metadata
        """
        for file_path in ppt_path_list:
            self.ppt_file_to_dict(file_path)

    def read_files_to_dict(self):
        """
        Reads each ppt file and saves the data and metadata to self.occur_dict
        """
        ppt_path_list = self.get_ppt_path_list(self.ppt_dir_path)
        self.ppt_files_to_dict(ppt_path_list)
