# open an excel spreadsheet

# later this can be done using user input and a path

# cycle through each column

# check content against criteria, using logic

# logic including regex checking for forbidden characters, certain characters, whitespace
# check for length

# IMPORTING LIBRARIES
import openpyxl
from pathlib import Path
import re
from openpyxl.styles import Font, PatternFill

# SETUP
path = Path.home()/'Desktop'/'template_dummy.xlsx'
error_fill =PatternFill(start_color="FF0000", end_color="FF0000", patternType="solid")
alphanumeric_regex = "[A-Z]"

# OPENING THE WORKBOOK AND TARGETING THE WORKSHEET
wb = openpyxl.load_workbook(path)
ws = wb["Catalogue Template"]

# CHECKING THE DEPARTMENT CODE COLUMN (D)
for col in ws.iter_cols(min_row=4, max_col=1):
    for cell in col:
        if not bool(re.match(alphanumeric_regex, str(cell.value))):
            cell.fill = error_fill

# CHECKING THE SERIES NUMBER
# these should all be the same number
# perhaps there could be some kind of user input to state what number this ought to be
# for now i will hard code it
series_num = 28

# SAVING THE CHANGES
wb.save(path)