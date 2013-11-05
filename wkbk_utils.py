#!/usr/bin/env python

#===============================================================================
#
# wkbk_utils.py - Set of utilities to help in using workbooks
#                   from xlrd and Xlsxwriter.
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
import argparse
import os
import os.path
import sys

import xlrd

import xlsxwriter

#=====
#
# CONSTANTS
#
#=====

#=====
#
# FUNCTIONS
#
#=====


#=====
#
# SCRIPT BODY
#
#=====

parser = argparse.ArgumentParser()

parser.add_argument("-v", "--verbose", action="store_true",
                    default=False,
                    help='Get real chatty [default = false]')

opts = parser.parse_args()


