#!/usr/bin/env python3

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
import argparse
import os.path
import sys

#import xlrd
import openpyxl

# Simulate an "include billing_common.py".
SCRIPT_DIR = os.path.dirname(os.path.realpath(__file__))
exec(compile(open(os.path.join(SCRIPT_DIR, "billing_common.py"), "rb").read(), os.path.join(SCRIPT_DIR, "billing_common.py"), 'exec'))

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
global argparse_get_parent_parser
global argparse_get_year_month
global argparse_get_billingroot_billingconfig

#=====
#
# SCRIPT BODY
#
#=====

parser = argparse.ArgumentParser(parents=[argparse_get_parent_parser()])

parser.add_argument("-a", "--accounting_file",
                    default=None,
                    help='The SGE accounting file to snapshot [default = None]')

args = parser.parse_args()

#
# Sanity-check arguments.
#

# Get year/month-related arguments
(year, month, begin_month_timestamp, end_month_timestamp) = argparse_get_year_month(args)

# Get BillingRoot and BillingConfig arguments
(billing_root, billing_config_file) = argparse_get_billingroot_billingconfig(args)

#
# Use values for accounting_file and billing_root from options, if available.
#   Open the BillingConfig file as a xlrd Workbook, if not.
#
accounting_file = args.accounting_file
if args.billing_config_file is not None and accounting_file is None:

    # Get absolute path for billing_config_file.
    billing_config_file = os.path.abspath(args.billing_config_file)

    #billing_config_wkbk = xlrd.open_workbook(billing_config_file)
    billing_config_wkbk = openpyxl.load_workbook(billing_config_file)
    config_dict = config_sheet_get_dict(billing_config_wkbk)

    accounting_file = config_dict.get("SGEAccountingFile")

if accounting_file is None:
    parser.exit(-1, "Need accounting file from BillingConfig file or command line switch.")

# Get absolute path for accounting_file
accounting_file = os.path.abspath(accounting_file)

# Within BillingRoot, create YEAR/MONTH dirs if necessary.
year_month_dir = os.path.join(billing_root, str(year), "%02d" % month)
if not os.path.exists(year_month_dir):
    os.makedirs(year_month_dir)

#
# Print summary of arguments.
#
print("TAKING ACCOUNTING FILE SNAPSHOT OF %02d/%d:" % (month, year))
print("  BillingConfigFile: %s" % billing_config_file)
print("  BillingRoot: %s" % billing_root)
print("  SGEAccountingFile: %s" % accounting_file)

# Create output accounting pathname.
new_accounting_filename = "%s.%d-%02d.txt" % (SGEACCOUNTING_PREFIX, year, month)
new_accounting_pathname = os.path.join(year_month_dir, new_accounting_filename)

print()
print("  OutputAccountingFile: %s" % new_accounting_pathname)

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

print()

accounting_output_fp.close()
accounting_input_fp.close()

print("Jobs found for %02d/%d:\t\t%d" % (month, year, this_months_job_count))
print("Total jobs in accounting:\t%d" % job_count)
