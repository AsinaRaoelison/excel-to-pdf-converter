# Excel to PDF Converter

This program converts non-empty ranges from an Excel (.xls*) file to PDF format. It accepts two command-line arguments: -f (--file) for the Excel file path and -w (--worksheet) for the worksheet name.

## Requirements

- Python 3.x
- pandas library (install using `pip install pandas`)
- reportlab library (install using `pip install reportlab`)

## Installation

1. Clone the repository to your local machine:

    git clone https://github.com/your_username/excel-to-pdf-converter.git
    cd excel-to-pdf-converter

2. Install the required libraries:

    pip install pandas reportlab

## Usage

To execute the program, use the following command:
    python excel-to-pdf-converter.py -f /path/to/excel_file.xls -w worksheet_name
    Replace `/path/to/excel_file.xls` with the path to your Excel file and `worksheet_name` with the name of the worksheet containing the non-empty range you want to convert.

    The program will generate a PDF file for each non-empty range found in the worksheet. The PDF files will be named as `output_1.pdf`, `output_2.pdf`, etc.

    If the provided arguments do not detect any non-empty range, an error message will be displayed.

## Example

    Suppose you have an Excel file named `data.xls` with a worksheet named `Sheet1` that contains a non-empty range. To convert this range to PDF, use the following command:

    python excel-to-pdf-converter.py -f data.xls -w Sheet1

    The program will create a PDF file named `output_1.pdf` containing the non-empty range from `Sheet1`.






