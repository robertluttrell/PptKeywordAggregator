class KeywordFilePresence:

    def __init__(self, file_path):
        self.num_instances = 0
        self.index_list = []
        self.file_path = file_path

    def __eq__(self, other):
        return self.num_instances == other.num_instances \
            and self.index_list == other.index_list \
            and self.file_path == other.file_path

    def add_index(self, index):
        self.index_list.append(index)
        self.num_instances += 1
