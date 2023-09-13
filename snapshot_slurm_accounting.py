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
import tempfile

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
                             "--delimiter=%s" % SLURMACCOUNTING_DELIMITER]

#=====
#
# FUNCTIONS
#
#=====
# From billing_common.py
global argparse_get_parent_parser
global argparse_get_year_month
global argparse_get_billingroot_billingconfig

def get_slurm_accounting(slurm_output_pathname, slurm_field_switches, filter_returns=False):

    slurm_accounting_file = open(slurm_output_pathname, "w")

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
    # COMMAND LINE: sacct SACCT_SWITCHES | fgrep -v PENDING | tr -d \r
    #

    (slurm_this_month_temp_file, slurm_this_month_temp_filename) = \
        tempfile.mkstemp(".txt", "%s-thisMonth." % temp_filename_prefix)

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

    fgrep_command_list = ['fgrep', '-v', '%sPENDING%s' % (SLURMACCOUNTING_DELIMITER, SLURMACCOUNTING_DELIMITER)]

    if filter_returns:
        fgrep_process = subprocess.Popen(fgrep_command_list,
                                         stdin=sacct_process.stdout,
                                         stdout=subprocess.PIPE)

        # Removing return characters embedded within lines
        tr_command_list = ['tr', '-d', "\r"]
        tr_process = subprocess.Popen(tr_command_list,
                                      stdin=fgrep_process.stdout,
                                      stdout=slurm_this_month_temp_file)

        (tr_process_stdout, tr_process_stderr) = tr_process.communicate()
    else:
        fgrep_process = subprocess.Popen(fgrep_command_list,
                                         stdin=sacct_process.stdout,
                                         stdout=slurm_this_month_temp_file)

    (fgrep_process_stdout, fgrep_process_stderr) = fgrep_process.communicate()
    (sacct_process_stdout, sacct_process_stderr) = sacct_process.communicate()

    #
    # "2. Get Slurm account for <MN+1>/01/<YR> through now (no End date needed)."
    #
    # COMMAND LINE: sacct SACCT_SWITCHES | fgrep -v PENDING | tr -d \r
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

    sacct_process = subprocess.Popen(slurm_command_list, stdout=subprocess.PIPE)

    fgrep_command_list = ['fgrep', '-v', '%sPENDING%s' % (SLURMACCOUNTING_DELIMITER, SLURMACCOUNTING_DELIMITER)]

    if filter_returns:
        fgrep_process = subprocess.Popen(fgrep_command_list,
                                         stdin=sacct_process.stdout,
                                         stdout=subprocess.PIPE)

        # Removing return characters embedded within lines
        tr_command_list = ['tr', '-d', "\r"]
        tr_process = subprocess.Popen(tr_command_list,
                                      stdin=fgrep_process.stdout,
                                      stdout=slurm_next_month_temp_file)

        (tr_process_stdout, tr_process_stderr) = tr_process.communicate()
    else:
        fgrep_process = subprocess.Popen(fgrep_command_list,
                                         stdin=sacct_process.stdout,
                                         stdout=slurm_next_month_temp_file)

    (fgrep_process_stdout, fgrep_process_stderr) = fgrep_process.communicate()
    (sacct_process_stdout, sacct_process_stderr) = sacct_process.communicate()

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

parser = argparse.ArgumentParser(parents=[argparse_get_parent_parser()])

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

# Get year/month-related arguments
(year, month, begin_month_timestamp, end_month_timestamp) = argparse_get_year_month(args)

# What month is it today?
today = datetime.date.today()
todays_month = today.month
todays_year  = today.year

# Calculate next month for range of this month.
if month != 12:
    next_month = month + 1
    next_month_year = year
else:
    next_month = 1
    next_month_year = year + 1

# Get BillingRoot and BillingConfig arguments
(billing_root, billing_config_file) = argparse_get_billingroot_billingconfig(args, year, month)

# Within BillingRoot, create YEAR/MONTH dirs if necessary.
year_month_dir = os.path.join(billing_root, str(year), "%02d" % month)
if not os.path.exists(year_month_dir):
    os.makedirs(year_month_dir)

#
# Print summary of arguments.
#
print("TAKING SLURM ACCOUNTING FILE SNAPSHOT OF %02d/%d:" % (month, year))
print("  BillingConfigFile: %s" % billing_config_file)
print("  BillingRoot: %s" % billing_root)

# Create output accounting pathnames.
slurm_accounting_filename_all = "%s.%d-%02d.all.txt" % (SLURMACCOUNTING_PREFIX, year, month)
slurm_accounting_pathname_all = os.path.join(year_month_dir, slurm_accounting_filename_all)

slurm_accounting_filename_min = "%s.%d-%02d.txt" % (SLURMACCOUNTING_PREFIX, year, month)
slurm_accounting_pathname_min = os.path.join(year_month_dir, slurm_accounting_filename_min)

print()
if not args.min_only: print("  SlurmAccountingFile (all): %s" % slurm_accounting_pathname_all)
if not args.all_only: print("  SlurmAccountingFile (min): %s" % slurm_accounting_pathname_min)
print()

if not args.min_only:
    print("GETTING SLURM ACCOUNTING - ALL FIELDS")
    get_slurm_accounting(slurm_accounting_pathname_all, SLURM_ACCT_FIELDS_ALL_SWITCHES, filter_returns=False)
else:
    print("**SKIPPING** SLURM ACCOUNTING - ALL FIELDS")
print()

if not args.all_only:
    print("GETTING SLURM ACCOUNTING - MINIMUM FIELDS")
    get_slurm_accounting(slurm_accounting_pathname_min, SLURM_ACCT_FIELDS_MIN_SWITCHES, filter_returns=True)
else:
    print("**SKIPPING** SLURM ACCOUNTING - MINIMUM FIELDS")
print()

