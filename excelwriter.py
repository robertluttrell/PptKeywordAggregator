class ExcelWriter:

    def __init__(self, excel_path):
        self.excel_path = excel_path

    def write_excel(self, df):
        """
        Writes the data in df to the excel sheet at self.excel_path
        :param df: DataFrame containing data about keywords and frequency
        """
        pass