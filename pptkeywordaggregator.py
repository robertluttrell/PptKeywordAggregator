import pptreader
import excelwriter


class PptKeywordAggregator:

    def __init__(self, args):
        self.ppt_dir_path = args.ppt_dir_path
        self.excel_path = args.excel_path

        self.ppt_reader = pptreader.PptReader(args.ppt_dir_path)
        self.excel_writer = excelwriter.ExcelWriter(args.excel_path)

    def run_program(self):
        self.ppt_reader.read_files_to_df()
