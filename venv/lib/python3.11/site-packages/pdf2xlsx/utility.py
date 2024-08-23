# -*- coding: utf-8 -*-
"""
Collection of utility functions
"""

def list2row(worksheet, row, col, values, positions=None):
    """
    Create header of the template xlsx file
        :param Worksheet worksheet: Worksheet class to write info
        :param int row: Row number to start writing
        :param int col: Column number to start writing
        :param list values: List of values to write in a row
        :param list positions: Positions for each value (otpional, if not given the
            values will be printed after each other from column 0)

        :return: the next position of cursor row,col
        :rtype: tuple of (int,int)
    """
    if not positions or len(positions) != len(values):
        positions = range(len(values))
    for val, pos in zip(values, positions):
        #worksheet.write(row, col+pos, val)
        worksheet.cell(row=row+1, column=col+pos+1, value=val)
    return row+1, col
