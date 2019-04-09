# Awards Parser

### Overview

The MAX Awards team recieves a list of nominations in a spreadsheet. Based on criterias for each award, the parser is able to parse through these submissions and create neat and readable word documents per nomination.


### Running it locally

NOTE: Ensure you have PIP and Python3 installed

1. `pip install docx python-docx openpyxl`
2. Place the nominations excel file in the same directory and ensure it is named: "2019 MAX Awards Nominations.xlsx"
3. Place the criteria excel file in the same directory and ensure it is named: "Awards_2019_Parsing.xlsx"
4. Run the parser `python3 awards_parser.py`

### Things to improve

1. Enable the program to accept command line arguments for:
   - the nominations file name but defaults to "2019 MAX Awards Nominations.xlsx" if not provided
   - the criteria file name but defaults to "Awards_2019_Parsing.xlsx" if not provided

### Known issues

1. Sometimes the award name in the criteria file and the nomination file don't match. Make sure that they do to get the complete list of the output