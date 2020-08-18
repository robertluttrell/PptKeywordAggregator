import glob
import os
import pandas as pd


class PptReader:

    def __init__(self, ppt_dir_path):
        self.ppt_dir_path = ppt_dir_path
        self.occur_dict = None

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

    @staticmethod
    def ppt_files_to_dict(ppt_path_list):
        """
        For each file specified in ppt_path_list, reads the file and stores keyword data in new dict
        :param ppt_path_list: List of strings specifying the powerpoint files to be processed
        :return: df: DataFrame containing keyword data and metadata
        """
        occur_dict = {}
        return occur_dict

    def read_files_to_dict(self):
        """
        Reads each ppt file and saves the data and metadata to self.occur_dict
        """
        ppt_path_list = self.get_ppt_path_list(self.ppt_dir_path)
        self.occur_dict = self.ppt_files_to_dict(ppt_path_list)
