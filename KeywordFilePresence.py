class KeywordFilePresence:

    def __init__(self, file_path):
        self.num_instances = 0
        self.index_list = []
        self.file_path = file_path

    def add_index(self, index):
        self.index_list.append(index)
        self.num_instances += 1

    def get_file_path(self):
        return self.file_path