import pptreader
import excelwriter


class PptKeywordAggregator:

    def __init__(self):
        self.ppt_reader = pptreader.PptReader()
        self.excel_writer = excelwriter.ExcelWriter()
