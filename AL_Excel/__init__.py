""" AL_Excel

    A fork of openpyxl with more of an emphasis on Object-Oriented manipulation
"""
## Super Module
from openpyxl import *
import openpyxl.worksheet.table, openpyxl.worksheet.worksheet
from openpyxl import utils

from AL_Excel.Coordinates import *
from AL_Excel.Ranges import *
from AL_Excel.Tables import *

## A keyfactory for removing spaces and lowercasing Headers for EnhancedTable.todicts()
lowerstrip = lambda key: key.lower().strip().replace(" ","_")


def maxcolumnwidth(worksheet, column, start_row = None, end_row = None):
    """ Determines the max length of the given column and returns it. This function ignores merged cells.
    
        worksheet should be a Worksheet instance.
        column should be a number or a column letter.
        start_row is the number of the row to start checking.
        end_row is the number of the row to finish checking (inclusive).
    """
    if isinstance(column, str):
        column = utils.column_index_from_string(column)

    start_row = start_row or 1
    end_row = end_row or worksheet.max_row
    if end_row <= start_row:
        raise ValueError("end_row should be greater or equal than start_row")

    _max = 0
    for row in range(start_row, end_row+1):
        cell = worksheet.cell(row, column)
        if cell.coordinate in worksheet.merged_cells:
            continue
        v = str(cell.value) if cell.value is not None else ""
        _max = max(_max, len(v))

    return _max



def autowidth(worksheet, column, start_row = None, end_row = None, padding = 0):
    """ Sets the width of a given column to it's max width (via maxcolumnwidth).
    
        worksheet should be a Worksheet instance
        column should be a number or a column letter.
        start_row is the number of the row to start checking.
        end_row is the number of the row to finish checking (inclusive).
        padding should be an integer which will be added to the width.
    """

    width = maxcolumnwidth(worksheet = worksheet, column = column, start_row= start_row, end_row= end_row)
    width += padding
    worksheet.column_dimensions[column].width = width