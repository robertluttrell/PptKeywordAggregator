import pptkeywordaggregator
import argparse


def parse_args():
    """
    Parses command-line arguments
    :return: argparse.Namespace object containing arguments
    """
    parser = argparse.ArgumentParser()
    parser.add_argument("ppt_dir_path", type=str, help="path to the directory containing ppt files")
    parser.add_argument("excel_path", type=str, help="path to the excel file")
    return parser.parse_args()


def main():
    args = parse_args()
    aggregator = pptkeywordaggregator.PptKeywordAggregator(args)
    aggregator.run_program()


if __name__ == "__main__":
    main()
