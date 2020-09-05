import pptkeywordaggregator
import argparse
import os


def parse_args():
    """
    Parses command-line arguments
    :return: argparse.Namespace object containing arguments
    """
    parser = argparse.ArgumentParser()

    parser.add_argument("ppt_dir_path", type=str, help="path to the directory containing ppt files")
    parser.add_argument("excel_path", type=str, help="path to the excel file")
    parser.add_argument("-gui", action="store_true", help="GUI flag")
    return parser.parse_args()


def get_ppt_path_list_cli(ppt_dir_path):
    """
    Creates a list of paths to all the powerpoint files in self.ppt_dir_path
    :return: list of paths
    """
    ppt_path_list = []
    cur_dir = os.getcwd()

    os.chdir(ppt_dir_path)
    for file in os.listdir(ppt_dir_path):
        if file.endswith(".pptx"):
            ppt_path_list.append(os.path.join(ppt_dir_path, file))

    os.chdir(cur_dir)
    return ppt_path_list


def main():
    args = parse_args()
    ppt_path_list = get_ppt_path_list_cli(args.ppt_dir_path)
    aggregator = pptkeywordaggregator.PptKeywordAggregator(ppt_path_list, args.excel_path)
    aggregator.run_program()


if __name__ == "__main__":
    main()
