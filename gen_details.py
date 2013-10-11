#!/usr/bin/env python

#===============================================================================
#
# gen_details.py - Generate billing details for month/year.
#
# ARGS:
#   1st: the BillingConfig spreadsheet.
#
# SWITCHES:
#
# OUTPUT:
#
# ASSUMPTIONS:
#   The input spreadsheet has been certified by check_config.py.
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
import datetime
import os
import os.path
import subprocess
import sys
import time

import xlrd
import xlsxwriter

#=====
#
# CONSTANTS
#
#=====
BILLING_DETAILS_PREFIX = "BillingDetails"

BILLING_DETAILS_SHEET_COLUMNS = {
    'Storage'    : ['Date', 'Folder', 'PI Tag', 'Size', 'Used'],
    'Computing'  : ['Date', 'User', 'Job Name', 'Account', 'Node', 'Cores', 'Wallclock'],
    'Consulting' : ['Date', 'Project', 'PI', 'Hours']
}

QUOTA_EXECUTABLE = ['/usr/lpp/mmfs/bin/mmlsquota', '-j']
USAGE_EXECUTABLE = ['du', '-s']
STORAGE_BLOCK_SIZE_ARG = ['--block-size=1G']  # Works in both above commands.

#
# Formats for the output workbook, to be initialized along with the workbook.
BOLD_FORMAT = None
DATE_FORMAT = None
FLOAT_FORMAT = None

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


def config_sheet_get_dict(wkbk):

    config_sheet = wkbk.sheet_by_name("Config")

    config_keys   = sheet_get_named_column(config_sheet, "Key")
    config_values = sheet_get_named_column(config_sheet, "Value")

    return dict(zip(config_keys, config_values))


def read_config_sheet(wkbk):

    config_dict = config_sheet_get_dict(wkbk)

    accounting_file = config_dict.get("SGEAccountingFile")
    if accounting_file is None:
        print >> sys.stderr, "Need accounting file: exiting..."
        sys.exit(-1)

    billing_root    = config_dict.get("BillingRoot", os.getcwd())

    return (billing_root, accounting_file)


def init_billing_details_wkbk(workbook):

    global BOLD_FORMAT
    global DATE_FORMAT
    global FLOAT_FORMAT

    # Create formats for use within the workbook
    BOLD_FORMAT = workbook.add_format({'bold' : True})
    DATE_FORMAT = workbook.add_format({'num_format' : 'mm/dd/yy'})
    FLOAT_FORMAT = workbook.add_format({'num_format' : '0.0'})

    sheet_name_to_sheet = dict()

    for sheet_name in BILLING_DETAILS_SHEET_COLUMNS.keys():

        sheet = workbook.add_worksheet(sheet_name)
        for col in range(0, len(BILLING_DETAILS_SHEET_COLUMNS[sheet_name])):
            sheet.write(0, col, BILLING_DETAILS_SHEET_COLUMNS[sheet_name][col], BOLD_FORMAT)

        sheet_name_to_sheet[sheet_name] = sheet

    return sheet_name_to_sheet


def get_folder_quota(folder, pi_tag):

    # Build and execute the quota command.
    quota_cmd = QUOTA_EXECUTABLE + ["projects." + pi_tag] + STORAGE_BLOCK_SIZE_ARG + ["gsfs0"]

    try:
        quota_output = subprocess.check_output(quota_cmd)
    except subprocess.CalledProcessError as cpe:
        print >> sys.stderr, "Couldn't get quota for %s (exit %d)" % (pi_tag, cpe.returncode)
        print >> sys.stderr, " Output: %s" % (quota_output)
        return None

    # Parse the results.
    for line in quota_output.split('\n'):

        fields = line.split()

        # If the first word on this line is 'gsfs0', this is the line we want.
        if fields[0] == 'gsfs0':
            used  = int(fields[2])
            quota = int(fields[3])

            return (used/1024.0, quota/1024.0)  # Return values in fractions of Tb.

    return None


def get_folder_usage(folder, pi_tag):

    # Build and execute the quota command.
    usage_cmd = USAGE_EXECUTABLE + STORAGE_BLOCK_SIZE_ARG + [folder]

    try:
        usage_output = subprocess.check_output(usage_cmd)
    except subprocess.CalledProcessError as cpe:
        print >> sys.stderr, "Couldn't get usage for %s (exit %d)" % (folder, cpe.returncode)
        print >> sys.stderr, " Output: %s" % (usage_output)
        return None

    # Parse the results.
    for line in usage_output.split('\n'):

        fields = line.split()
        used = fields[0]

        return (used/1024, used/1024)  # Return values in fractions of Tb.

    return None


def compute_storage_charges(config_wkbk, begin_timestamp, end_timestamp, storage_sheet):

    # Get lists of folders, quota booleans.
    folder_sheet = config_wkbk.sheet_by_name('Folders')

    folders     = sheet_get_named_column(folder_sheet, 'Folder')
    pi_tags     = sheet_get_named_column(folder_sheet, 'PI Tag')
    quota_bools = sheet_get_named_column(folder_sheet, 'By Quota?')

    folder_sizes = []
    # Create mapping from folders to space used.
    for (folder, pi_tag, quota_bool) in zip(folders, pi_tags, quota_bools):

        if quota_bool == "yes":
            # Check folder's quota.
            (used, total) = get_folder_quota(folder, pi_tag)
        else:
            # Check folder's usage.
            (used, total) = get_folder_usage(folder, pi_tag)

        folder_sizes.append([ folder, pi_tag, total, used ])

    #
    # Write space used mapping into details workbook.
    #
    for row in range(0,len(folder_sizes)):

        # 'Date'
        col = 0
        storage_sheet.write_formula(row + 1, col, '=%f/(60*60*24)+DATE(1970,1,1)' % time.time(), DATE_FORMAT)
        print time.time(),

        # 'Folder'
        col += 1
        storage_sheet.write(row + 1, col, folder_sizes[row][0])
        print folder_sizes[row][0],

        # 'PI Tag'
        col += 1
        storage_sheet.write(row + 1, col, folder_sizes[row][1])
        print folder_sizes[row][1],

        # 'Size'
        col += 1
        storage_sheet.write(row + 1, col, folder_sizes[row][2], FLOAT_FORMAT)
        print folder_sizes[row][2],

        # 'Used'
        col += 1
        storage_sheet.write(row + 1, col, folder_sizes[row][3], FLOAT_FORMAT)
        print folder_sizes[row][3]

def compute_computing_charges(config_wkbk, begin_timestamp, end_timestamp, accounting_file, computing_sheet):
    pass

def compute_consulting_charges(config_wkbk, begin_timestamp, end_timestamp, consulting_sheet):
    pass

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
parser.add_argument("-y","--year", type=int, choices=range(2013,2031),
                    default=None,
                    help="The year to be filtered out. [default = this year]")
parser.add_argument("-m", "--month", type=int, choices=range(1,13),
                    default=None,
                    help="The month to be filtered out. [default = last month]")

args = parser.parse_args()

billing_config_wkbk = xlrd.open_workbook(args.billing_config_file)

#
# Process arguments.
#

# Do year first, because month might modify it.
if args.year is None:
    year = datetime.date.today().year
else:
    year = args.year

# Do month now, and decrement year if want last month and this month is Dec.
if args.month is None:
    # No month given: use last month.
    this_month = datetime.date.today().month

    # If this month is Jan, last month was Dec. of previous year.
    if this_month == 1:
        month = 12
        year -= 1
    else:
        month = this_month - 1
else:
    month = args.month

# Calculate next month for range of this month.
if month != 12:
    next_month = month + 1
    next_month_year = year
else:
    next_month = 1
    next_month_year = year + 1

# The begin_ and end_month_timestamps are to be used as follows:
#   date is within the month if begin_month_timestamp <= date < end_month_timestamp
# Both values should be GMT.
begin_month_timestamp = int(time.mktime(datetime.date(year, month, 1).timetuple()))
end_month_timestamp   = int(time.mktime(datetime.date(next_month_year, next_month, 1).timetuple()))

#
# Get the location of the BillingRoot directory from the Config sheet.
# Get the location of the accounting file from the Config sheet.
#
(billing_root, accounting_file) = read_config_sheet(billing_config_wkbk)

# Initialize the BillingDetails spreadsheet.
details_wkbk_filename = "%s.%s-%02d.xlsx" % (BILLING_DETAILS_PREFIX, year, month)
details_wkbk_pathname = os.path.join(billing_root, details_wkbk_filename)

billing_details_wkbk = xlsxwriter.Workbook(details_wkbk_pathname)
sheet_name_to_sheet = init_billing_details_wkbk(billing_details_wkbk)

#
# Compute storage charges.
#
compute_storage_charges(billing_config_wkbk, begin_month_timestamp, end_month_timestamp, sheet_name_to_sheet['Storage'])

#
# Compute computing charges.
#
compute_computing_charges(billing_config_wkbk, begin_month_timestamp, end_month_timestamp, accounting_file, sheet_name_to_sheet['Computing'])

#
# Compute consulting charges.
#
compute_consulting_charges(billing_config_wkbk, begin_month_timestamp, end_month_timestamp, sheet_name_to_sheet['Consulting'])

#
# Close the output workbook and write the .xlsx file.
#
billing_details_wkbk.close()