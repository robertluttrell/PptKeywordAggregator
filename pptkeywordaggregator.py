import pptreader
import excelwriter


class PptKeywordAggregator:

    def __init__(self, args):
        self.ppt_dir_path = args.ppt_dir_path
        self.excel_path = args.excel_path

        self.ppt_reader = pptreader.PptReader(args.ppt_dir_path)
        self.excel_writer = None

    def run_program(self):
        ppt_path_list = pptreader.PptReader.get_ppt_path_list(self.ppt_dir_path)
        self.ppt_reader.ppt_files_to_dict(ppt_path_list)

        self.excel_writer = excelwriter.ExcelWriter(self.excel_path, self.ppt_reader.word_dict)
        self.excel_writer.write_excel()


