#!/usr/bin/env python

#===============================================================================
#
# NAMEOFSCRIPT.py - Description of script
#
# ARGS:
#
# SWITCHES:
#
# OUTPUT:
#
# ASSUMPTIONS:
#
# AUTHOR:
#   Keith Bettinger
#
#==============================================================================

#=====
#
# IMPORTS
#
#=====
from optparse import OptionParser
import os
import os.path
import sys

import xlrd

class ConfigWorkbook(xlrd.Book):

#=====
#
# CONSTANTS
#
#=====

# Constants for sheet names.
# Constants for column names.

#=====
#
# FUNCTIONS
#
#=====

  # This method takes in an xlrd Sheet object and a column name,
  # and returns all the values from that column.
  def sheet_get_named_column(sheet, col_name):

    header_row = sheet.row_values(0)

    for idx in range(len(header_row)):
        if header_row[idx] == col_name:
           col_name_idx = idx
           break
    else:
        return None

    return sheet.col_values(col_name_idx,start_rowx=1)

  def open_workbook(self, wkbk_file):
      return xlrd.open_workbook(wkbk_file)



