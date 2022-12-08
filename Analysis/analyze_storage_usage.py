#!/usr/bin/env python

# ===============================================================================
#
# analyze_storage_usage.py - Description of script
#
# ARGS:
#  All: BillingDetails files
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
# ==============================================================================

#=====
#
# IMPORTS
#
#=====
import argparse
from collections import defaultdict
import datetime
import os
import os.path
import xlrd
import xlsxwriter
import sys

# Simulate an "include billing_common.py".
SCRIPT_DIR = os.path.dirname(os.path.realpath(__file__))
exec(compile(open(os.path.join(SCRIPT_DIR, "..", "billing_common.py"), "rb").read(), os.path.join(SCRIPT_DIR, "..", "billing_common.py"), 'exec'))

global from_excel_date_to_date_string

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
def read_billingdetails_file(bd_wkbk):

    global folder_to_date_size
    global folder_to_date_used
    global folder_set
    global date_set

    storage_sheet = bd_wkbk.sheet_by_name("Storage")

    for row in range(1,storage_sheet.nrows):

        (date, timestamp, folder, size, used) = storage_sheet.row_values(row)

        measured_date = datetime.datetime.utcfromtimestamp(timestamp)
        # Move date to last day of previous month.
        measured_date -= datetime.timedelta(measured_date.day + 1)

        ym_date = measured_date.strftime("%Y/%m")

        folder_to_date_size[(folder,ym_date)] = size
        folder_to_date_used[(folder,ym_date)] = used

        folder_set.add(folder)
        date_set.add(ym_date)


def write_storage_wkbk(out_wkbk, sheet_name, folder_to_date_dict):

    sheet = out_wkbk.add_worksheet(sheet_name)

    # Write the first column of folder names.
    sheet.write_column(1,0, sorted(folder_set))

    col = 1
    for date in sorted(date_set):

        date_col = [date]
        for folder in sorted(folder_set):

            if (folder,date) in folder_to_date_dict:
                date_col.append(folder_to_date_dict[(folder,date)])
            else:
                date_col.append('N/A')

        sheet.write_column(0,col, date_col)
        col += 1


#=====
#
# SCRIPT BODY
#
#=====

parser = argparse.ArgumentParser()

parser.add_argument("billing_details_files", nargs="+")
parser.add_argument("-o","--output",
                    default="storage_history.xlsx",
                    help='Filename of resulting .xlsx file.')
parser.add_argument("-v", "--verbose", action="store_true",
                    default=False,
                    help='Get chatty [default = false]')
parser.add_argument("-d", "--debug", action="store_true",
                    default=False,
                    help='Get REAL chatty [default = false]')

args = parser.parse_args()

# Initialize data structures.
folder_to_date_size = dict()  # From (folder, date) to size
folder_to_date_used = dict()  # From (folder, date) to used

folder_set = set()  # Set of all folders (rows of output sheets)
date_set   = set()  # Set of all dates   (cols of output sheets)

for billing_details_file in args.billing_details_files:

    if billing_details_file.endswith('.xlsx'):
        # Open the BillingDetails workbook.
        print("Reading BillingDetails workbook %s." % billing_details_file)
        billing_details_wkbk = xlrd.open_workbook(billing_details_file, on_demand=True)

        read_billingdetails_file(billing_details_wkbk)

        billing_details_wkbk.release_resources()

print("Writing out storage workbook %s" % args.output)
storage_wkbk = xlsxwriter.Workbook(args.output)

write_storage_wkbk(storage_wkbk, 'Quotas', folder_to_date_size)
write_storage_wkbk(storage_wkbk, 'Usages', folder_to_date_used)

storage_wkbk.close()
