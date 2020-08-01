import pandas as pd


class PptReader:

    def __init__(self, ppt_path_list):
        self.ppt_path_list = ppt_path_list
        self.df = None

    @staticmethod
    def ppt_files_to_df(ppt_path_list):
        """
        For each file specified in ppt_path_list, reads the file and stores keyword data in a new DataFrame
        :param ppt_path_list: List of strings specifying the powerpoint files to be processed
        :return: df: DataFrame containing keyword data and metadata
        """
        df = pd.DataFrame()
        return df

    def process_files(self):
        self.df = self.ppt_files_to_df(self.ppt_path_list)

