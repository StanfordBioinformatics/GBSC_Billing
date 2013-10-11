#!/usr/bin/env python

#===============================================================================
#
# check_config.py - Confirm that the BillingConfiguration workbook makes sense.
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

#=====
#
# CONSTANTS
#
#=====
BILLING_CONFIG_SHEETS = ['Rates', 'PIs', 'Folders', 'Users', 'JobTags', 'Config']

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

    return sheet.col_values(col_name_idx, start_rowx=1)

def check_sheets(wkbk):

    any_errors = False

    all_sheets = wkbk.sheet_names()
    print all_sheets
    for sheet in BILLING_CONFIG_SHEETS:
        if sheet not in all_sheets:
            print "check_sheets:  Missing sheet %s" % sheet
            any_errors = True

    return any_errors

def check_pctages(sheet, col_name):

    any_errors = False

    objects = sheet_get_named_column(sheet, col_name)
    pctages = sheet_get_named_column(sheet, '%age')

    # Create a mapping from an object to the list of percentages associated with it.
    object_pctages = dict()
    for (fdr, pct) in zip(objects, pctages):
        if fdr not in object_pctages.keys():
            object_pctages[fdr] = [pct]
        else:
            object_pctages[fdr].append(pct)

    # Add all the percentages for each object, and confirm that the sum is 100.
    for object in object_pctages.keys():
        pctage_list = map(float, object_pctages[object])
        total_pctage = sum(pctage_list)
        if total_pctage != 1.0:
            print "check_pctages: %s %s percentages add up to %d%%, not 100%%" % (col_name, object, total_pctage*100.0)
            any_errors = True

    return any_errors

#=====
#
# SCRIPT BODY
#
#=====

parser = argparse.ArgumentParser()

parser.add_argument("billing_config_file",
                    help='The BillingConfig file')
parser.add_argument("-v", "--verbose", action="store_true",
                    default=False,
                    help='Get real chatty [default = false]')

args = parser.parse_args()

billing_config_wkbk = xlrd.open_workbook(args.billing_config_file)

# Check that all the proper sheets are in
#  TODO: Maybe go into their structure?
sheets_are_OK = check_sheets(billing_config_wkbk)

# Check that, in the Folders sheet, the %ages associated with each folder
#  add up to 100%.
if billing_config_wkbk.sheet_loaded('Folders'):
    folder_sheet = billing_config_wkbk.sheet_by_name('Folders')
    folder_pctages_are_OK = check_pctages(folder_sheet, 'Folder')
else:
    folder_pctages_are_OK = False

# Check that, in the JobTags sheet, the %ages associated with each job tag
#  add up to 100%.
if billing_config_wkbk.sheet_loaded('JobTags'):
    jobtag_sheet = billing_config_wkbk.sheet_by_name('JobTags')
    jobtag_pctages_are_OK = check_pctages(jobtag_sheet, 'Job Tag')
else:
    jobtag_pctages_are_OK = False

if not (sheets_are_OK and folder_pctages_are_OK and jobtag_pctages_are_OK):
    print "Billing Configuration spreadsheet has some problems."
    sys.exit(-1)
else:
    sys.exit(0)