#!/usr/bin/env python

# ===============================================================================
#
# confirm_jobIDs.py - Confirms throughput of jobIDs in all stages of processing.
#
# ARGS:
#  1st: the BillingConfig spreadsheet.
#
# SWITCHES:
#   --accounting_file: Location of accounting file (overrides BillingConfig.xlsx)
#   --billing_details_file: Location of the BillingDetails.xlsx file (default=look in BillingRoot/<year>/<month>)
#   --billing_root:    Location of BillingRoot directory (overrides BillingConfig.xlsx)
#                      [default if no BillingRoot in BillingConfig.xlsx or switch given: CWD]
#   --year:            Year of snapshot requested. [Default is this year]
#   --month:           Month of snapshot requested. [Default is last month]
#
# OUTPUT:
#   Text to stdout regarding the flow of jobIDs from accounting file
#     to BillingDetails file to BillingNotifs files.
#
# ASSUMPTIONS:
#   BillingNotifs files corresponding to all the PI tags within the BillingConfig
#     file are in the BillingRoot directory.
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
from collections import defaultdict
import datetime
import os
import os.path
import sys
import xlrd

# Simulate an "include billing_common.py".
SCRIPT_DIR = os.path.dirname(os.path.realpath(__file__))
execfile(os.path.join(SCRIPT_DIR, "billing_common.py"))

#=====
#
# GLOBALS
#
#=====
# In billing_common.py
global SGEACCOUNTING_PREFIX
global BILLING_DETAILS_PREFIX
global BILLING_NOTIFS_PREFIX
global ACCOUNTING_FIELDS

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
# from billing_common.py
global read_config_sheet
global from_excel_date_to_timestamp
global from_excel_date_to_date_string
global from_ymd_date_to_timestamp
global sheet_get_named_column

def get_pi_tag_list(billing_config_wkbk):

    # Get PI tag list from BillingConfig workbook.
    pis_sheet   = billing_config_wkbk.sheet_by_name("PIs")
    pi_tag_list = sheet_get_named_column(pis_sheet, "PI Tag")

    #
    # Filter pi_tag_list for PIs active in the current month.
    #
    pi_dates_added   = sheet_get_named_column(pis_sheet, "Date Added")
    pi_dates_removed = sheet_get_named_column(pis_sheet, "Date Removed")

    pi_tags_and_dates_added = zip(pi_tag_list, pi_dates_added, pi_dates_removed)

    for (pi_tag, date_added, date_removed) in pi_tags_and_dates_added:

        # Convert the Excel dates to timestamps.
        date_added_timestamp = from_excel_date_to_timestamp(date_added)
        if date_removed != '':
            date_removed_timestamp = from_excel_date_to_timestamp(date_removed)
        else:
            date_removed_timestamp = end_month_timestamp + 1  # Not in this month.

        # If the date added is AFTER the end of this month, or
        #  the date removed is BEFORE the beginning of this month,
        # then remove the pi_tag from the list.
        if date_added_timestamp >= end_month_timestamp:

            print >> sys.stderr, " *** Ignoring PI %s: added after this month on %s" % (pi_tag, from_excel_date_to_date_string(date_added))
            pi_tag_list.remove(pi_tag)

        elif date_removed_timestamp < begin_month_timestamp:

            print >> sys.stderr, " *** Ignoring PI %s: removed before this month on %s" % (pi_tag, from_excel_date_to_date_string(date_removed))
            pi_tag_list.remove(pi_tag)

    return pi_tag_list


def read_jobIDs(wkbk, sheet_name):

    sheet  = wkbk.sheet_by_name(sheet_name)
    jobIDs = sheet_get_named_column(sheet, "Job ID")

    return map(lambda x: int(x), jobIDs)

def print_set(my_set, max_elts=10000):

    elt_count = 0
    for elt in my_set:
        if elt_count % 10 == 0:
            elt_count += 1
            print
        if elt_count >= max_elts:
            break

        print "%s" % elt,
    else:
        print


#=====
#
# SCRIPT BODY
#
#=====

parser = argparse.ArgumentParser()

parser.add_argument("billing_config_file",
                    help='The BillingConfig file')
parser.add_argument("-a", "--accounting_file",
                    default=None,
                    help='The SGE accounting file to read [default = None]')
parser.add_argument("-d","--billing_details_file",
                    default=None,
                    help='The BillingDetails file')
parser.add_argument("-r", "--billing_root",
                    default=None,
                    help='The Billing Root directory [default = None]')
parser.add_argument("-y","--year", type=int, choices=range(2013,2031),
                    default=None,
                    help="The year to be filtered out. [default = this year]")
parser.add_argument("-m", "--month", type=int, choices=range(1,13),
                    default=None,
                    help="The month to be filtered out. [default = last month]")

parser.add_argument("-v", "--verbose", action="store_true",
                    default=False,
                    help='Get chatty [default = false]')


args = parser.parse_args()

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
# Both values should be UTC.
begin_month_timestamp = from_ymd_date_to_timestamp(year, month, 1)
end_month_timestamp   = from_ymd_date_to_timestamp(next_month_year, next_month, 1)

#
# Open the Billing Config workbook.
#
billing_config_wkbk = xlrd.open_workbook(args.billing_config_file)

#
# Get the location of the BillingRoot directory from the Config sheet.
#  (Ignore the accounting file from this sheet).
#
(billing_root, _) = read_config_sheet(billing_config_wkbk)

# Override billing_root with switch args, if present.
if args.billing_root is not None:
    billing_root = args.billing_root
# If we still don't have a billing root dir, use the current directory.
if billing_root is None:
    billing_root = os.getcwd()

# Within BillingRoot, create YEAR/MONTH dirs if necessary.
year_month_dir = os.path.join(billing_root, str(year), "%02d" % month)

# Use switch arg for accounting_file if present, else use file in BillingRoot.
if args.accounting_file is not None:
    accounting_file = args.accounting_file
else:
    accounting_filename = "%s.%d-%02d.txt" % (SGEACCOUNTING_PREFIX, year, month)
    accounting_file = os.path.join(year_month_dir, accounting_filename)

# If BillingDetails file given, use that, else look in BillingRoot.
if args.billing_details_file is not None:
    billing_details_file = args.billing_details_file
else:
    billing_details_file = os.path.join(year_month_dir, "%s.%d-%02d.xlsx" % (BILLING_DETAILS_PREFIX, year, month))

#
# Read all JobIDs from accounting file.
#
accounting_jobIDs = []

accounting_file_fp = open(accounting_file)

for line in accounting_file_fp:

    if line[0] == "#": continue

    fields = line.split(':')
    fields_dict = dict(zip(ACCOUNTING_FIELDS,fields))

    accounting_jobIDs.append(int(fields_dict['job_number']))

accounting_file_fp.close()

#
# Read all JobIDs from BillingDetails file.
#  Sheets:
#   Computing (Billable Jobs)
#   Nonbillable Jobs
#   Failed Jobs
#
billing_details_wkbk = xlrd.open_workbook(billing_details_file)

details_billable_jobIDs    = read_jobIDs(billing_details_wkbk, 'Computing')
details_nonbillable_jobIDs = read_jobIDs(billing_details_wkbk, 'Nonbillable Jobs')
details_failed_jobIDs      = read_jobIDs(billing_details_wkbk, 'Failed Jobs')

#
# For each PI tag, read the BillingNotifs file.
#  Sheet: Computing Details
#

# Get list of active PI tags.
pi_tag_list = get_pi_tag_list(billing_config_wkbk)

# Make mapping from PI tag to list of jobIDs.
pi_tag_jobIDs_dict = defaultdict(list)

# Loop over PI tag list to get jobIDs
for pi_tag in pi_tag_list:

    notifs_wkbk_filename = "%s-%s.%s-%02d.xlsx" % (BILLING_NOTIFS_PREFIX, pi_tag, year, month)
    notifs_wkbk_pathname = os.path.join(year_month_dir, notifs_wkbk_filename)

    billing_notifs_wkbk = xlrd.open_workbook(notifs_wkbk_pathname)

    pi_tag_jobIDs = read_jobIDs(billing_notifs_wkbk, "Computing Details")
    pi_tag_jobIDs_dict[pi_tag].extend(pi_tag_jobIDs)


#
# Analyze the JobID sources.
#

# Unique the accounting JobIDs.
accounting_all_jobID_set = set(accounting_jobIDs)

# Aggregate the BillingDetails JobIDs and unique them.
details_all_jobID_set = set(reduce(lambda a,b: a+b,[details_billable_jobIDs,details_nonbillable_jobIDs,details_failed_jobIDs]))

# Unique the BillingDetails Billable JobIDs.
details_billable_jobID_set = set(details_billable_jobIDs)

# Aggregate the BillingNotifs JobIDs and unique them.
notifs_all_jobID_set  = set(reduce(lambda a,b: a+b, pi_tag_jobIDs_dict.values()))

# Compare:
#  The accounting JobIDs
#  The BillingDetails JobID aggregate
# They should be the same.

# This operation gets set of elements in either accounting or details, but not both.
accounting_symdiff_details = accounting_all_jobID_set ^ details_all_jobID_set

print "NOT IN BOTH ACCOUNTING AND DETAILS: %d" % (len(accounting_symdiff_details))
print

if len(accounting_symdiff_details) > 0:
    print "In accounting only:"
    print_set(accounting_symdiff_details & accounting_all_jobID_set, 100)
    print

    print "In details only:"
    print_set(accounting_symdiff_details & details_all_jobID_set, 100)
    print

# Compare:
#  The BillingDetails Billable JobIDs
#  The BillingNotifs JobID aggregate
# They should be the same.

# This operation gets set of elements in either billable details or notifs, but not both.
details_billable_symdiff_notifs = details_billable_jobID_set ^ notifs_all_jobID_set

print "NOT IN BILLABLE DETAILS AND NOTIFS: %d" % (len(details_billable_symdiff_notifs))
print

if len(details_billable_symdiff_notifs) > 0:
    print "In billable details only:"
    print_set(details_billable_symdiff_notifs & details_billable_jobID_set)
    print

    print "In notifs only:"
    print_set(details_billable_symdiff_notifs & notifs_all_jobID_set)
    print

if len(accounting_symdiff_details) > 0 or len(details_billable_symdiff_notifs) > 0:
    print
    print "JOBS INCONSISTENT AMONG FILES"

    sys.exit(-1)
else:
    print
    print "ALL JOBS ARE IN ALL FILES"

    sys.exit(0)