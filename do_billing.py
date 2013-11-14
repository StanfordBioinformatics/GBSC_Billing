#!/usr/bin/env python

#===============================================================================
#
# do_billing.py - Run all the scripts in order to create new BillingNotifications
#                  for a particular month/year.
#
# ARGS:
#   1st: the BillingConfig spreadsheet.
#
# SWITCHES:
#   --year:            Year of snapshot requested. [Default is this year]
#   --month:           Month of snapshot requested. [Default is last month]
#   --force:           Skip check on BillingConfig sheet. [default=False]
#   --no_storage:      Don't run the storage calculations. [default=Run them]
#   --no_usage:        Don't run the storage usage calculations (only the quotas).
#   --no_computing:    Don't run the computing calculations. [default=Run them]
#   --no_consulting:   Don't run the consulting calculations. [default=Run them].
#   --all_jobs_billable: Regard all jobs as billable. [default=False]
#   --pi_sheets:       Add PI-specific sheets to BillingAggregate workbook. [default=False]
#
# OUTPUT:
#   The output of each script run is written to STDOUT.
#
# ASSUMPTIONS:
#   The scripts to be executed by this one live in the directory with it.
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

#=====
#
# CONSTANTS
#
#=====
SCRIPT_DIR = os.path.dirname(os.path.realpath(__file__))  # The directory of this script.

# These scripts are assumed to live in SCRIPT_DIR.
SNAPSHOT_ACCT_SCRIPT = "snapshot_accounting.py"
CHECK_CONFIG_SCRIPT  = "check_config.py"
GEN_DETAILS_SCRIPT   = "gen_details.py"
GEN_NOTIFS_SCRIPT    = "gen_notifs.py"

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

parser.add_argument("billing_config_file",
                    default=None,
                    help='The BillingConfig file [default = None]')
parser.add_argument("-v", "--verbose", action="store_true",
                    default=False,
                    help='Get real chatty [default = false]')
parser.add_argument("-f", "--force", action="store_true",
                    default=False,
                    help='Skip check on config file [default = false]')
parser.add_argument("--no_storage", action="store_true",
                    default=False,
                    help="Don't run storage calculations [default = false]")
parser.add_argument("--no_usage", action="store_true",
                    default=False,
                    help="Don't run storage usage calculations [default = false]")
parser.add_argument("--no_computing", action="store_true",
                    default=False,
                    help="Don't run computing calculations [default = false]")
parser.add_argument("--no_consulting", action="store_true",
                    default=False,
                    help="Don't run consulting calculations [default = false]")
parser.add_argument("--all_jobs_billable", action="store_true",
                    default=False,
                    help="Consider all jobs to be billable [default = false]")
parser.add_argument("-p", "--pi_sheets", action="store_true",
                    default=False,
                    help='Add PI-specific sheets to BillingAggregate file [default = false]')
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

# Save year and month arguments, which appear in almost every command.
year_month_args = ['-y', str(year), '-m', str(month)]

###
#
# Run the check_config.py on the BillingConfig file.
#
###
if not args.force:
    check_config_script_path = os.path.join(SCRIPT_DIR, CHECK_CONFIG_SCRIPT)
    check_config_cmd = [check_config_script_path, args.billing_config_file]

    print 'RUNNING CHECK CONFIG:'
    if args.verbose: print check_config_cmd
    try:
        check_config_output = subprocess.check_output(check_config_cmd, stderr=subprocess.STDOUT)
    except subprocess.CalledProcessError as cpe:
        print >> sys.stderr, "Check config on %s failed (exit %d)" % (args.billing_config_file, cpe.returncode)
        print >> sys.stderr, " Output: %s" % (cpe.output)
        sys.exit(-1)

    print check_config_output
    print

###
#
# Snapshot the accounting file.
#
###
snapshot_script_path = os.path.join(SCRIPT_DIR, SNAPSHOT_ACCT_SCRIPT)
snapshot_cmd = [snapshot_script_path] + year_month_args + ['-c', args.billing_config_file]

print "RUNNING SNAPSHOT ACCOUNTING:"
if args.verbose: print snapshot_cmd
try:
    subprocess.check_call(snapshot_cmd, stderr=subprocess.STDOUT)
except subprocess.CalledProcessError as cpe:
    print >> sys.stderr, "Snapshot accounting on %s failed (exit %d)" % (args.billing_config_file, cpe.returncode)
    print >> sys.stderr, " Output: %s" % (cpe.output)
    sys.exit(-1)

print

###
#
# Generate the BillingDetails file.
#
###
details_script_path = os.path.join(SCRIPT_DIR, GEN_DETAILS_SCRIPT)
details_cmd = [details_script_path] + year_month_args + [args.billing_config_file]

# Add the switches concerning details.
if args.no_storage:        details_cmd += ['--no_storage']
if args.no_usage:          details_cmd += ['--no_usage']
if args.no_computing:      details_cmd += ['--no_computing']
if args.no_consulting:     details_cmd += ['--no_consulting']
if args.all_jobs_billable: details_cmd += ['--all_jobs_billable']

print "RUNNING GENERATE DETAILS:"
if args.verbose: print details_cmd
try:
    subprocess.check_call(details_cmd, stderr=subprocess.STDOUT)
except subprocess.CalledProcessError as cpe:
    print >> sys.stderr, "Generate Details on %s failed (exit %d)" % (args.billing_config_file, cpe.returncode)
    print >> sys.stderr, " Output: %s" % (cpe.output)
    sys.exit(-1)

print

###
#
# Generate the BillingNotifications files.
#
###
notifs_script_path = os.path.join(SCRIPT_DIR, GEN_NOTIFS_SCRIPT)
notifs_cmd = [notifs_script_path] + year_month_args + [args.billing_config_file]

# Add the --pi_sheets switch, if requested.
if args.pi_sheets: notifs_cmd += ['--pi_sheets']

print "RUNNING GENERATE NOTIFICATIONS:"
if args.verbose: print notifs_cmd
try:
    subprocess.check_call(notifs_cmd, stderr=subprocess.STDOUT)
except subprocess.CalledProcessError as cpe:
    print >> sys.stderr, "Generate Notifications on %s failed (exit %d)" % (args.billing_config_file, cpe.returncode)
    print >> sys.stderr, " Output: %s" % (cpe.output)
    sys.exit(-1)

print

print "BILLING SCRIPTS COMPLETE."