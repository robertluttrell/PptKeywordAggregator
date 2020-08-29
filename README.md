# PptKeywordAggregator

## Overview
PptKeywordAggregator reads keywords from PowerPoint presentations and organizes them in an Excel spreadsheet.

## Features
- Reads keywords from text-containing elements that start with the prefix "Keywords:" on each slide of each presentation selected
- Writes data to Excel spreadsheet including:
    - Keyword name
    - Filename it was found in (one filename per Excel row entry)
    - Count of occurrences of the keyword in the specified Number of occurrences of the word in the specified presentation
    - Indices of each slide the keyword was found in for the specific presentation
- Supports both Windows and Unix-based operating systems (MacOS and Linux)
    
## Running the Project
1. Clone this project locally
2. Run `pip install python-pptx` in your bash/commandline
3. Run `pip install openpyxl` in your bash/commandline
4. Run the program using
```
python3 main.py <ppt_dir_path> <excel_file_path>
```
The program will process all files ending in `.pptx` in `<ppt_dir_path>` and save the keyword data to a new spreadsheet at `<excel_file_path>`.

**Note: any existing file matching `<excel_file_path>` will be overwritten.**

## Dependencies
- Python 3
- [python-pptx](https://python-pptx.readthedocs.io/en/latest/)
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/)
