import pptkeywordaggregator
import argparse


def parse_args():
    """
    Parses command-line arguments
    :return: argparse.Namespace object containing arguments
    """
    parser = argparse.ArgumentParser()

    parser.add_argument("-gui", action="store_true", help="GUI flag")
    parser.add_argument("excel_path", type=str, help="path to the excel file")
    parser.add_argument("ppt_paths", nargs='+')
    return parser.parse_args()


def main():
    args = parse_args()
    ppt_path_list = args.ppt_paths
    aggregator = pptkeywordaggregator.PptKeywordAggregator(ppt_path_list, args.excel_path)
    aggregator.run_program()


if __name__ == "__main__":
    main()
