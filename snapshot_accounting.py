#!/usr/bin/env python

#===============================================================================
#
# snapshot_accounting.py - Copies the given month/year's SGE accounting data
#                           into a separate file.
#
# ARGS:
#   1st: BillingConfig.xlsx file (for Config sheet: location of accounting file)
#        [optional if --accounting_file given]
#
# SWITCHES:
#   --accounting_file: Location of accounting file (overrides BillingConfig.xlsx)
#   --billing_root:    Location of BillingRoot directory (overrides BillingConfig.xlsx)
#                      [default if no BillingRoot in BillingConfig.xlsx or switch given: CWD]
#   --billing_config_file: Location of BillingConfig xlsx file
#                      [not required if both --accounting_file and --billing_root given].
#   --year:            Year of snapshot requested. [Default is this year]
#   --month:           Month of snapshot requested. [Default is last month]
#
# OUTPUT:
#    An accounting file with only entries with end_dates within the given
#     month.  This file, named SGEAccounting.<YEAR>-<MONTH>.txt, will be placed in
#     <BillingRoot>/<YEAR>/<MONTH>/ if BillingRoot is given or in the current
#     working directory if not.
#
# ASSUMPTIONS:
#    Dependent on xlrd module.
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

import datetime
import time
import argparse
import os
import os.path
import sys

import xlrd

# Simulate an "include billing_common.py".
SCRIPT_DIR = os.path.dirname(os.path.realpath(__file__))
execfile(os.path.join(SCRIPT_DIR, "billing_common.py"))

#=====
#
# CONSTANTS
#
#=====
# From billing_common.py
global SGEACCOUNTING_PREFIX
global ACCOUNTING_FAILED_CODES

#=====
#
# FUNCTIONS
#
#=====
# From billing_common.py
global config_sheet_get_dict
global from_ymd_date_to_timestamp

#=====
#
# SCRIPT BODY
#
#=====

parser = argparse.ArgumentParser()

parser.add_argument("-a", "--accounting_file",
                    default=None,
                    help='The SGE accounting file to snapshot [default = None]')
parser.add_argument("-r", "--billing_root",
                    default=None,
                    help='The Billing Root directory [default = None]')
parser.add_argument("-c", "--billing_config_file",
                    default=None,
                    help='The BillingConfig file [default = None]')
parser.add_argument("-v", "--verbose", action="store_true",
                    default=False,
                    help='Get real chatty [default = false]')
parser.add_argument("-y","--year", type=int, choices=range(2013,2021),
                    default=None,
                    help="The year to be filtered out. [default = this year]")
parser.add_argument("-m", "--month", type=int, choices=range(1,13),
                    default=None,
                    help="The month to be filtered out. [default = last month]")

args = parser.parse_args()

#
# Sanity-check arguments.
#

# If there is no billing_config_file and no accounting_file,
# flag an error.
if args.billing_config_file is None and args.accounting_file is None:
    parser.print_usage()
    parser.exit(-1, "Need either --billing_config_file or --accounting_file.\n")

# Do year next, because month might modify it.
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
# Use values for accounting_file and billing_root from options, if available.
#   Open the BillingConfig file as a xlrd Workbook, if not.
#
if args.billing_config_file is not None:

    # Get absolute path for billing_config_file.
    billing_config_file = os.path.abspath(args.billing_config_file)

    billing_config_wkbk = xlrd.open_workbook(billing_config_file)
    config_dict = config_sheet_get_dict(billing_config_wkbk)

    accounting_file = config_dict.get("SGEAccountingFile")
    billing_root    = config_dict.get("BillingRoot")
else:
    billing_config_file = None
    accounting_file = None
    billing_root    = None

# Override billing_root with switch args, if present.
if args.billing_root is not None:
    billing_root = args.billing_root
# If we still don't have a billing root dir, use the current directory.
if billing_root is None:
    billing_root = os.getcwd()

# Get absolute path for billing_root
billing_root = os.path.abspath(billing_root)

# Within BillingRoot, create YEAR/MONTH dirs if necessary.
year_month_dir = os.path.join(billing_root, str(year), "%02d" % month)
if not os.path.exists(year_month_dir):
    os.makedirs(year_month_dir)

# Use switch arg for accounting_file if present.
if args.accounting_file is not None:
    accounting_file = args.accounting_file

# Get absolute path for accounting_file
accounting_file = os.path.abspath(accounting_file)

if accounting_file is None:
    parser.exit(-1, "Need accounting file from BillingConfig file or command line switch.")

#
# Print summary of arguments.
#
print "TAKING ACCOUNTING FILE SNAPSHOT OF %02d/%d:" % (month, year)
print "  BillingConfigFile: %s" % (billing_config_file)
print "  BillingRoot: %s" % (billing_root)
print "  SGEAccountingFile: %s" % (accounting_file)

# Create output accounting pathname.
new_accounting_filename = "%s.%d-%02d.txt" % (SGEACCOUNTING_PREFIX, year, month)
new_accounting_pathname = os.path.join(year_month_dir, new_accounting_filename)

print
print "  OutputAccountingFile: %s" % (new_accounting_pathname)

#
# Open the current accounting file for input.
#
accounting_input_fp = open(accounting_file, "r")

#
# Open the new accounting file for output.
#
accounting_output_fp = open(new_accounting_pathname, "w")

#
# Read all the lines of the current accounting file.
#  Output to the new accounting file all those lines
#  which have "end_dates" in the given month.
#
job_count = 0
this_months_job_count = 0
for line in accounting_input_fp:

    if line[0] == "#": continue

    fields = line.split(':')
    submission_date = int(fields[8])
    end_date = int(fields[10])
    failed = int(fields[11])

    # If this job failed, then use its submission_time as the job date.
    # else use the end_time as the job date.
    job_failed = failed in ACCOUNTING_FAILED_CODES
    if job_failed:
        job_date = submission_date  # No end_date for failed jobs.
    else:
        job_date = end_date

    # If the end date of this job was within the month,
    #  output it to the new accounting file.
    found_job = False
    if begin_month_timestamp <= job_date < end_month_timestamp:
        accounting_output_fp.write(line)
        this_months_job_count += 1
        found_job = True

    if job_count % 10000 == 0:
        if found_job:
            sys.stdout.write(':')
        else:
            sys.stdout.write('.')
        sys.stdout.flush()
    job_count += 1

print

accounting_output_fp.close()
accounting_input_fp.close()

print "Jobs found for %02d/%d:\t\t%d" % (month, year, this_months_job_count)
print "Total jobs in accounting:\t%d" % (job_count)
