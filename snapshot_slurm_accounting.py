#!/usr/bin/env python3

#===============================================================================
#
# snapshot_slurm_accounting.py - Copies the given month/year's Slurm accounting data
#                                 into a separate file.
#
# ARGS:
#   1st: BillingConfig.xlsx file
#
# SWITCHES:
#   --billing_root:    Location of BillingRoot directory (overrides BillingConfig.xlsx)
#                      [default if no BillingRoot in BillingConfig.xlsx or switch given: current working dir]
#   --year:            Year of snapshot requested. [Default is this year]
#   --month:           Month of snapshot requested. [Default is last month]
#
# OUTPUT:
#    An Slurm accounting file with only entries with end_dates within the given
#     month.  This file, named SlurmAccounting.<YEAR>-<MONTH>.txt, will be placed in
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
import subprocess
import sys
import tempfile

import xlrd

# for SLURMACCOUNTING_DELIMITER
from slurm_job_accounting_entry import SlurmJobAccountingEntry

# Simulate an "include billing_common.py".
SCRIPT_DIR = os.path.dirname(os.path.realpath(__file__))
exec(compile(open(os.path.join(SCRIPT_DIR, "billing_common.py"), "rb").read(), os.path.join(SCRIPT_DIR, "billing_common.py"), 'exec'))

#=====
#
# CONSTANTS
#
#=====
# From billing_common.py
global SLURMACCOUNTING_PREFIX
global SLURMACCOUNTING_DELIMITER

SLURM_ACCT_COMMAND_NAME = ["sacct"]
SLURM_ACCT_STATE_SWITCHES = ["--state=CA,CD,DL,F,NF,PR,TO,OOM,RQ"]

# We will collect both all the fields from Slurm records
# and just the ones we need for reports within our system.
SLURM_ACCT_FIELDS_ALL_SWITCHES = ["--format=ALL"]
SLURM_ACCT_FIELDS_MIN_SWITCHES = ["--format=User,JobName,Account,WCKey,NodeList,NCPUS,ElapsedRaw,JobID,JobIDRaw,MaxVMSize,Submit,Start,End"]

SLURM_ACCT_OTHER_SWITCHES = ["--allusers","--parsable2","--allocations","--duplicates",
                             "--delimiter=%s" % (SlurmJobAccountingEntry.DELIMITER_HASH)]

#=====
#
# FUNCTIONS
#
#=====
# From billing_common.py
global config_sheet_get_dict

def get_slurm_accounting(output_pathname, slurm_field_switches):

    slurm_accounting_file = open(output_pathname, "w")

    ###
    #
    # Steps to get Slurm accounting for only this month:
    #   1. Get Slurm accounting for <MN>/01/<YR> through <MN+1>/01/<YR>.
    #       - This data includes jobs running during the month but ending after, so...
    #   2. Get Slurm account for <MN+1>/01/<YR> through now (no End date needed).
    #   3. Subtract the lines from 2. from the lines in 1.
    #
    ###

    temp_filename_prefix = "%s.%d-%02d" % (SLURMACCOUNTING_PREFIX, year, month)

    #
    # "1. Get Slurm accounting for <MN>/01/<YR> through <MN+1>/01/<YR>."
    #

    (slurm_this_month_temp_file, slurm_this_month_temp_filename) = tempfile.mkstemp(".txt",
                                                                                    "%s-thisMonth." % temp_filename_prefix)

    print("Getting Slurm accounting for %d-%02d to %s" % (year, month, slurm_this_month_temp_filename))

    # Create the start and end dates switches to the command.
    slurm_command_starttime_switch = ["--starttime", "%02d/01/%02d" % (month, year - 2000)]
    # If we are in the month we are querying for, use the current date
    if month == todays_month and year == todays_year:
        slurm_command_endtime_switch = ["--endtime", datetime.datetime.today().strftime("%m/%d/%y-%H:%M:%S")]
    else:
        slurm_command_endtime_switch = ["--endtime", "%02d/01/%02d" % (next_month, next_month_year - 2000)]

    slurm_command_list = SLURM_ACCT_COMMAND_NAME + SLURM_ACCT_STATE_SWITCHES + \
                         slurm_field_switches + \
                         SLURM_ACCT_OTHER_SWITCHES + \
                         slurm_command_starttime_switch + slurm_command_endtime_switch

    if args.verbose:
        print(slurm_command_list)

    sacct_process = subprocess.Popen(slurm_command_list, stdout=subprocess.PIPE)

    fgrep_command_list = ['fgrep', '-v', '|PENDING|']

    fgrep_process = subprocess.Popen(fgrep_command_list,
                                     stdin=sacct_process.stdout,
                                     stdout=slurm_this_month_temp_file)

    (fgrep_process_stdout, fgrep_process_stderr) = fgrep_process.communicate()
    (sacct_process_stdout, sacct_process_stderr) = sacct_process.communicate()

    #
    # "2. Get Slurm account for <MN+1>/01/<YR> through now (no End date needed)."
    #

    (slurm_next_month_temp_file, slurm_next_month_temp_filename) = tempfile.mkstemp(".txt",
                                                                                    "%s-nextMonth." % temp_filename_prefix)

    print("Getting Slurm accounting for %d-%02d to %s" % (next_month_year, next_month, slurm_next_month_temp_filename))

    # Create the start date switch to the command.
    slurm_command_starttime_switch = ["--starttime", "%02d/01/%02d" % (next_month, next_month_year - 2000)]

    slurm_command_list = SLURM_ACCT_COMMAND_NAME + SLURM_ACCT_STATE_SWITCHES + \
                         slurm_field_switches + \
                         SLURM_ACCT_OTHER_SWITCHES + \
                         slurm_command_starttime_switch + ["--noheader"]

    if args.verbose:
        print(slurm_command_list)

    ret_val = subprocess.call(slurm_command_list,
                              stdout=slurm_next_month_temp_file)

    #
    # 3. Subtract the lines from 2. from the lines in 1.
    #
    print("Subtracting the two files")

    ret_val = subprocess.call(["awk", "{if (f==1) { r[$0] } else if (! ($0 in r)) { print $0 } }",
                               "f=1", slurm_next_month_temp_filename, "f=2", slurm_this_month_temp_filename],
                              stdout=slurm_accounting_file)


#=====
#
# SCRIPT BODY
#
#=====

parser = argparse.ArgumentParser()

parser.add_argument("billing_config_file",
                    default=None,
                    help='The BillingConfig file [default = None]')
parser.add_argument("-r", "--billing_root",
                    default=None,
                    help='The Billing Root directory [default = None]')
parser.add_argument("-v", "--verbose", action="store_true",
                    default=False,
                    help='Get real chatty [default = false]')
parser.add_argument("-y","--year", type=int, choices=list(range(2013,2040)),
                    default=None,
                    help="The year to be filtered out. [default = this year]")
parser.add_argument("-m", "--month", type=int, choices=list(range(1,13)),
                    default=None,
                    help="The month to be filtered out. [default = last month]")
parser.add_argument("-a", "--all_only", action="store_true",
                    default=False,
                    help='Only output the complete accounting file [default = output both all and min]')
parser.add_argument("-n", "--min_only", action="store_true",
                    default=False,
                    help='Only output the minimum accounting file [default = output both all and min]')

args = parser.parse_args()

#
# Sanity-check arguments.
#

# Do year next, because month might modify it.
if args.year is None:
    year = datetime.date.today().year
else:
    year = args.year

# What month is it today?
today = datetime.date.today()
todays_month = today.month
todays_year  = today.year

# No month given: use last month.
#  Do month now, and decrement year if want last month and this month is Dec.
if args.month is None:

    # If this month is Jan, last month was Dec. of previous year.
    if todays_month == 1:
        month = 12
        year -= 1
    else:
        month = todays_month - 1
else:
    month = args.month

# Calculate next month for range of this month.
if month != 12:
    next_month = month + 1
    next_month_year = year
else:
    next_month = 1
    next_month_year = year + 1

#
# Use value for billing_root from switches, if available.
#   Open the BillingConfig file as a xlrd Workbook to find BillingRoot, if not.
#
if args.billing_config_file is not None:

    # Get absolute path for billing_config_file.
    billing_config_file = os.path.abspath(args.billing_config_file)

    billing_config_wkbk = xlrd.open_workbook(billing_config_file)
    config_dict = config_sheet_get_dict(billing_config_wkbk)

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

#
# Print summary of arguments.
#
print("TAKING SLURM ACCOUNTING FILE SNAPSHOT OF %02d/%d:" % (month, year))
print("  BillingConfigFile: %s" % (billing_config_file))
print("  BillingRoot: %s" % (billing_root))

# Create output accounting pathnames.
slurm_accounting_filename_all = "%s.%d-%02d.all.txt" % (SLURMACCOUNTING_PREFIX, year, month)
slurm_accounting_pathname_all = os.path.join(year_month_dir, slurm_accounting_filename_all)

slurm_accounting_filename_min = "%s.%d-%02d.txt" % (SLURMACCOUNTING_PREFIX, year, month)
slurm_accounting_pathname_min = os.path.join(year_month_dir, slurm_accounting_filename_min)

print()
if not args.min_only: print("  SlurmAccountingFile (all): %s" % (slurm_accounting_pathname_all))
if not args.all_only: print("  SlurmAccountingFile (min): %s" % (slurm_accounting_pathname_min))
print()

if not args.min_only:
    print("GETTING SLURM ACCOUNTING - ALL FIELDS")
    get_slurm_accounting(slurm_accounting_pathname_all, SLURM_ACCT_FIELDS_ALL_SWITCHES)
else:
    print("**SKIPPING** SLURM ACCOUNTING - ALL FIELDS")
print()

if not args.all_only:
    print("GETTING SLURM ACCOUNTING - MINIMUM FIELDS")
    get_slurm_accounting(slurm_accounting_pathname_min, SLURM_ACCT_FIELDS_MIN_SWITCHES)
else:
    print("**SKIPPING** SLURM ACCOUNTING - MINIMUM FIELDS")
print()

