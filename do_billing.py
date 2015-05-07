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
import shutil
import stat
import subprocess
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

# These scripts are assumed to live in SCRIPT_DIR.
SNAPSHOT_ACCT_SCRIPT = "snapshot_accounting.py"
CHECK_CONFIG_SCRIPT  = "check_config.py"
GEN_DETAILS_SCRIPT   = "gen_details.py"
GEN_NOTIFS_SCRIPT    = "gen_notifs.py"
ILAB_EXPORT_SCRIPT   = "gen_ilab_upload.py"

ILAB_AVAILABLE_SERVICES_FILE = "available_services_list_150106.csv"
ILAB_EXPORT_TEMPLATE_FILE    = "charges_upload_template.csv"

#=====
#
# FUNCTIONS
#
#=====
# From billing_common.py
global read_config_sheet

#=====
#
# SCRIPT BODY
#
#=====

parser = argparse.ArgumentParser()

parser.add_argument("--billing_config_file",
                    default=None,
                    help='The BillingConfig file [default = None]')
parser.add_argument("--billing_root",
                    default=None,
                    help='The BillingRoot directory [default = None]')
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

# Make little directory hierarchy for year and month.
year_month_dir = os.path.join(str(year), "%02d" % month)

# Find BillingConfig file and BillingRoot dir if not given.
if args.billing_config_file is None:

    if args.billing_root is None:
        # ERROR: Need billing_config_file OR billing_root.
        parser.exit(-1, "Need either billing_config_file or billing_root\n")

    else:
        # Use billing_root given as argument.
        billing_root = args.billing_root

        # Look for billing_config_file in given billing_root dir.
        billing_config_file = os.path.join(billing_root,"BillingConfig.xlsx")

else:
    # Use billing_config_file given as argument.
    billing_config_file = args.billing_config_file

    # Was billing_root given as argument?
    if args.billing_root is None:
        # Look in billing_config_file for billing_root.
        config_wkbk = xlrd.open_workbook(billing_config_file)
        (billing_root, _) = read_config_sheet(config_wkbk)
    else:
        # Use billing_root given as argument.
        billing_root = args.billing_root


if not os.path.exists(billing_config_file):
    # ERROR: Can't find billing_config_file
    parser.exit(-2, "Can't find billing config file %s\n" % billing_config_file)

if not os.path.exists(billing_root):
    # ERROR: Can't find billing_root
    parser.exit(-3, "Can't find billing root dir %s\n" % billing_root)

# Add the year/month dir hierarchy to billing_root.
year_month_root = os.path.join(billing_root, year_month_dir)
if not os.path.exists(year_month_root):
    os.makedirs(year_month_root, 0770)

# Copy billing config file into year_month_root, unless they are the same file.
billing_config_file_copy = os.path.join(billing_root,year_month_dir,"BillingConfig.%s-%02d.xlsx" % (year,month))
if billing_config_file != billing_config_file_copy:
    shutil.copyfile(billing_config_file, billing_config_file_copy)

# Save year and month arguments, which appear in almost every command.
year_month_args = ['-y', str(year), '-m', str(month)]
# Save billing_root argument, now used in every command.
billing_root_args = ['--billing_root', billing_root]

#
# Open file for output for all scripts into BillingRoot dir.
#
billing_log_file = open(os.path.join(year_month_root,"BillingLog.%d-%02d.txt" % (year,month)), 'w')

###
#
# Run the check_config.py on the BillingConfig file.
#
###
if not args.force:
    check_config_script_path = os.path.join(SCRIPT_DIR, CHECK_CONFIG_SCRIPT)
    check_config_cmd = [check_config_script_path, billing_config_file]

    print 'CHECKING BILLING CONFIG:'
    print >> billing_log_file, 'CHECKING BILLING CONFIG:'
    if args.verbose: print >> billing_log_file, check_config_cmd
    try:
        subprocess.check_call(check_config_cmd, stdout=billing_log_file, stderr=subprocess.STDOUT)
    except subprocess.CalledProcessError as cpe:
        print >> sys.stderr, "Check config on %s failed (exit %d)" % (billing_config_file, cpe.returncode)
        print >> sys.stderr, " Output: %s" % (cpe.output)
        sys.exit(-1)

    print
    print >> billing_log_file

###
#
# Snapshot the accounting file.
#
###
snapshot_script_path = os.path.join(SCRIPT_DIR, SNAPSHOT_ACCT_SCRIPT)
snapshot_cmd = [snapshot_script_path] + year_month_args + billing_root_args + ['-c', billing_config_file]

print "SNAPSHOTTING ACCOUNTING:"
print >> billing_log_file, "SNAPSHOTTING ACCOUNTING:"
if args.verbose: print >> billing_log_file, snapshot_cmd
try:
    subprocess.check_call(snapshot_cmd, stdout=billing_log_file, stderr=subprocess.STDOUT)
except subprocess.CalledProcessError as cpe:
    print >> sys.stderr, "Snapshot accounting on %s failed (exit %d)" % (billing_config_file, cpe.returncode)
    print >> sys.stderr, " Output: %s" % (cpe.output)
    sys.exit(-1)

print
print >> billing_log_file

###
#
# Generate the BillingDetails file.
#
###
details_script_path = os.path.join(SCRIPT_DIR, GEN_DETAILS_SCRIPT)
details_cmd = [details_script_path] + year_month_args + billing_root_args + [billing_config_file]

# Add the switches concerning details.
if args.no_storage:        details_cmd += ['--no_storage']
if args.no_usage:          details_cmd += ['--no_usage']
if args.no_computing:      details_cmd += ['--no_computing']
if args.no_consulting:     details_cmd += ['--no_consulting']
if args.all_jobs_billable: details_cmd += ['--all_jobs_billable']

print "GENERATING DETAILS:"
print >> billing_log_file, "GENERATING DETAILS:"
if args.verbose: print >> billing_log_file, details_cmd
try:
    subprocess.check_call(details_cmd, stdout=billing_log_file, stderr=subprocess.STDOUT)
except subprocess.CalledProcessError as cpe:
    print >> sys.stderr, "Generate Details on %s failed (exit %d)" % (billing_config_file, cpe.returncode)
    print >> sys.stderr, " Output: %s" % (cpe.output)
    sys.exit(-1)

print
print >> billing_log_file

###
#
# Generate the BillingNotifications files.
#
###
notifs_script_path = os.path.join(SCRIPT_DIR, GEN_NOTIFS_SCRIPT)
notifs_cmd = [notifs_script_path] + year_month_args + billing_root_args + [billing_config_file]

# Add the --pi_sheets switch, if requested.
if args.pi_sheets: notifs_cmd += ['--pi_sheets']

print "GENERATING NOTIFICATIONS:"
print >> billing_log_file, "GENERATING NOTIFICATIONS:"
if args.verbose: print >> billing_log_file, notifs_cmd
try:
    notifs_output = subprocess.check_call(notifs_cmd, stdout=billing_log_file, stderr=subprocess.STDOUT)
except subprocess.CalledProcessError as cpe:
    print >> sys.stderr, "Generate Notifications on %s failed (exit %d)" % (billing_config_file, cpe.returncode)
    print >> sys.stderr, " Output: %s" % (cpe.output)
    sys.exit(-1)

print
print >> billing_log_file

###
#
# Generate the BillingiLab file.
#
###
ilab_export_script_path = os.path.join(SCRIPT_DIR, ILAB_EXPORT_SCRIPT)
ilab_available_services_path = os.path.join(SCRIPT_DIR, 'iLab', ILAB_AVAILABLE_SERVICES_FILE)
ilab_export_template_path    = os.path.join(SCRIPT_DIR, 'iLab', ILAB_EXPORT_TEMPLATE_FILE)

ilab_export_cmd = [ilab_export_script_path] + year_month_args + billing_root_args + [billing_config_file]

ilab_export_cmd += ['--ilab_template', ilab_export_template_path]
ilab_export_cmd += ['--ilab_available_services', ilab_available_services_path]

print "EXPORTING TO ILAB:"
print >> billing_log_file, "EXPORTING TO ILAB:"
if args.verbose: print >> billing_log_file, ilab_export_cmd
try:
    ilab_export_output = subprocess.check_call(ilab_export_cmd, stdout=billing_log_file, stderr=subprocess.STDOUT)
except subprocess.CalledProcessError as cpe:
    print >> sys.stderr, "iLab Export on %s failed (exit %d)" % (billing_config_file, cpe.returncode)
    print >> sys.stderr, " Output: %s" % (cpe.output)
    sys.exit(-1)

print
print >> billing_log_file

###
#
# Set all the files in the year/month dir to read-only.
#
###
# Permissions: User: rX, Group: rX, Other: none
dir_mode  = stat.S_IRUSR | stat.S_IXUSR | stat.S_IRGRP | stat.S_IXGRP
file_mode = stat.S_IRUSR | stat.S_IRGRP
for root, dirs, files in os.walk(year_month_root):
    os.chmod(root, dir_mode)
    for d in dirs:  os.chmod(os.path.join(root,d), dir_mode)
    for f in files: os.chmod(os.path.join(root,f), file_mode)

print "BILLING SCRIPTS COMPLETE."
print >> billing_log_file, "BILLING SCRIPTS COMPLETE."