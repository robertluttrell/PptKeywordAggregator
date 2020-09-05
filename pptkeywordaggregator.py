import pptreader
import excelwriter
import os


class PptKeywordAggregator:

    def __init__(self, ppt_path_list, excel_path):
        self.excel_path = excel_path
        self.ppt_path_list = ppt_path_list

        self.ppt_reader = pptreader.PptReader(ppt_path_list)
        self.excel_writer = None

    def get_ppt_path_list_gui(self):
        return []

    def run_program(self):
        self.ppt_reader = pptreader.PptReader(self.ppt_path_list)
        self.ppt_reader.ppt_files_to_dict()

        self.excel_writer = excelwriter.ExcelWriter(self.excel_path, self.ppt_reader.word_dict)
        self.excel_writer.write_excel()


