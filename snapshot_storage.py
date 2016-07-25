#!/usr/bin/env python

#===============================================================================
#
# snapshot_storage.py - Measure the usage and quotas for all the folders in the
#                         BillingConfig file and output them into a CSV file.
#
# ARGS:
#   1st: BillingConfig.xlsx file
#
# SWITCHES:
#     -h, --help            show this help message and exit
#     -r BILLING_ROOT, --billing_root BILLING_ROOT
#                         The Billing Root directory [default = None]
#     --no_usage            Don't run storage usage calculations [default = false]
#     -s STORAGE_USAGE_CSV, --storage_usage_csv STORAGE_USAGE_CSV
#                         The storage usage CSV file.
#     -y YEAR, --year YEAR
#                         The year to be filtered out. [default = this year]
#     -m {1,2,3,4,5,6,7,8,9,10,11,12}, --month {1,2,3,4,5,6,7,8,9,10,11,12}
#                         The month to be filtered out. [default = last month]
#     -v, --verbose         Get chatty [default = false]
#     -d, --debug           Get REAL chatty [default = false]
#
# OUTPUT:
#   <BillingRoot>/<year>/<month>/StorageUsage.<year>-<month>.csv file with
#     rows of the form Date, Timestamp, Folder, Size, Used.
#
# ASSUMPTIONS:
#   Script is run on a machine which can run the quota command.
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
import csv
import datetime
import os
import os.path
import sys
import subprocess
import time

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
global QUOTA_EXECUTABLE
global USAGE_EXECUTABLE
global STORAGE_BLOCK_SIZE_ARG
global STORAGE_PREFIX
global BILLING_DETAILS_SHEET_COLUMNS

#=====
#
# FUNCTIONS
#
#=====
# In billing_common.py
global sheet_get_named_column
global from_timestamp_to_excel_date
global from_excel_date_to_timestamp
global from_ymd_date_to_timestamp
global read_config_sheet

# Deduces the fileset that a folder being measured by quota lives in.
#
# Patterns:
#  /srv/gsfs0/DIR1      - Device "gsfs0", Fileset "DIR1"
#  /srv/gsfs0/DIR1/DIR2 - Device "gsfs0", Fileset "DIR1.DIR2"
#  /srv/gsfs0/BaaS/Labs/DIR1 - Device "gsfs0", Fileset "BaaS.DIR1"
#
def get_device_and_fileset_from_folder(folder):

    # We only know about "/srv/gsfs0" paths.
    if not folder.startswith("/srv/gsfs0"):
        print >> sys.stderr, "get_device_and_fileset_from_folder(): Cannot get device and fileset from %s: doesn't start with /srv/gsfs0" % folder
        return None

    # Find the two path elements after "/srv/gsfs0/".
    path_elts = os.path.normpath(folder).split(os.path.sep)

    # Expect at least ['', 'srv', 'gsfs0', DIR1].
    if len(path_elts) >= 4:

        # Pattern "/srv/gsfs0/BaaS/Labs/DIR1" ?
        if path_elts[3] == 'BaaS' and path_elts[4] == 'Labs':
            return (path_elts[2], ".".join([path_elts[3],path_elts[5]]))
        else:  # Other two patterns above.
            return (path_elts[2], ".".join(path_elts[3:5]))
    else:
        print >> sys.stderr, "get_device_and_fileset_from_folder(): Path %s is not long enough" % folder
        return None


# Gets the quota for the given PI tag.
# Returns a tuple of (size used, quota) with values in Tb, or
# None if there was a problem parsing the quota command output.
def get_folder_quota(machine, folder, pi_tag):

    if args.verbose: print "  Getting folder quota for %s..." % (pi_tag)

    # Build and execute the quota command.
    quota_cmd = []

    # Add ssh if this is a remote call.
    if machine is not None:
        quota_cmd += ["ssh", machine]

    # Find the fileset to get the quota of, from the folder name.
    device_and_fileset = get_device_and_fileset_from_folder(folder)
    if device_and_fileset is None:
        print >> sys.stderr, "ERROR: No fileset for folder %s; ignoring..." % (folder)
        return None

    (device, fileset) = device_and_fileset

    # Build the quota command from the assembled arguments.
    quota_cmd += QUOTA_EXECUTABLE + [fileset] + STORAGE_BLOCK_SIZE_ARG + [device]

    try:
        quota_output = subprocess.check_output(quota_cmd)
    except subprocess.CalledProcessError as cpe:
        print >> sys.stderr, "Couldn't get quota for %s (exit %d)" % (pi_tag, cpe.returncode)
        print >> sys.stderr, " Command:", quota_cmd
        print >> sys.stderr, " Output: %s" % (cpe.output)
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

# Gets the usage for the given folder.
# Returns a tuple of (quota, quota) with values in Tb, or
# None if there was a problem parsing the usage command output.
def get_folder_usage(machine, folder, pi_tag):

    if args.verbose: print "  Getting folder usage of %s..." % (folder)

    # Build and execute the quota command.
    usage_cmd = []

    # Add ssh if this is a remote call.
    if machine is not None:
        usage_cmd += ["ssh", machine]
    else:
        # Local du's require sudo.
        usage_cmd += ["sudo"]

    usage_cmd += USAGE_EXECUTABLE + STORAGE_BLOCK_SIZE_ARG + [folder]

    try:
        usage_output = subprocess.check_output(usage_cmd)
    except subprocess.CalledProcessError as cpe:
        print >> sys.stderr, "Couldn't get usage for %s (exit %d)" % (folder, cpe.returncode)
        print >> sys.stderr, " Command:", usage_cmd
        print >> sys.stderr, " Output: %s" % (cpe.output)
        return None

    # Parse the results.
    for line in usage_output.split('\n'):

        fields = line.split()
        used = int(fields[0])

        return (used/1024.0, used/1024.0)  # Return values in fractions of Tb.

    return None


# Generates the Storage sheet data from folder quotas and usages.
# Returns mapping from folders to [timestamp, total, used]
def compute_storage_charges(config_wkbk, begin_timestamp, end_timestamp):

    print "Computing storage charges..."

    # Lists of folders to measure come from:
    #  "PI Folder" column of "PIs" sheet, and
    #  "Folder" column of "Folders" sheet.

    # Get lists of folders, quota booleans from PIs sheet.
    pis_sheet = config_wkbk.sheet_by_name('PIs')

    folders     = sheet_get_named_column(pis_sheet, 'PI Folder')
    pi_tags     = sheet_get_named_column(pis_sheet, 'PI Tag')
    quota_bools = ['quota'] * len(folders)   # All PI folders are measured by quota.
    dates_added = sheet_get_named_column(pis_sheet, 'Date Added')
    dates_remvd = sheet_get_named_column(pis_sheet, 'Date Removed')

    # Get lists of folders, quota booleans from Folders sheet.
    folder_sheet = config_wkbk.sheet_by_name('Folders')

    folders     += sheet_get_named_column(folder_sheet, 'Folder')
    pi_tags     += sheet_get_named_column(folder_sheet, 'PI Tag')
    quota_bools += sheet_get_named_column(folder_sheet, 'Method')
    dates_added += sheet_get_named_column(folder_sheet, 'Date Added')
    dates_remvd += sheet_get_named_column(folder_sheet, 'Date Removed')

    # List of dictionaries with keys from BILLING_DETAILS_SHEET_COLUMNS['Storage'].
    folder_size_dicts = []
    # Set of folders that have been measured
    folders_measured = set()

    # Create mapping from folders to space used.
    for (folder, pi_tag, quota_bool, date_added, date_removed) in zip(folders, pi_tags, quota_bools, dates_added, dates_remvd):

        # Skip measuring this folder entry if the folder is None.
        if folder == 'None': continue

        # Skip measuring this folder if we have already done it.
        if folder in folders_measured: continue

        # If this folder has been added prior to or within this month
        # and has not been removed before the beginning of this month, analyze it.
        if (end_timestamp > from_excel_date_to_timestamp(date_added) and
            (date_removed == '' or begin_timestamp < from_excel_date_to_timestamp(date_removed)) ):

            # Split folder into machine:dir components.
            if folder.find(':') >= 0:
                (machine, dir) = folder.split(':')
            else:
                machine = None
                dir = folder

            if quota_bool == "quota":
                # Check folder's quota.
                print "Getting quota for %s" % folder
                (used, total) = get_folder_quota(machine, dir, pi_tag)
            elif quota_bool == "usage" and not args.no_usage:
                # Check folder's usage.
                print "Measuring usage for %s" % folder
                (used, total) = get_folder_usage(machine, dir, pi_tag)
            else:
                # Use null values for no usage data.
                print "SKIPPING measurement for %s" % folder
                (used, total) = (0, 0)

            # Record when we measured the storage.
            measured_timestamp = time.time()

            folder_size_dicts.append({ 'Date Measured' : from_timestamp_to_excel_date(measured_timestamp),
                                       'Timestamp' : measured_timestamp,
                                       'Folder' : folder,
                                       'Size' : total,
                                       'Used' :used })

            # Add the folder measured to a set of folders we have measured.
            folders_measured.add(folder)

        else:
            print "  *** Excluding %s for PI %s: folder not active in this month" % (folder, pi_tag)

    return folder_size_dicts


# Write storage usage data into a storage usage CSV file.
# Takes in mapping from folders to [timestamp, total, used].
def write_storage_usage_data(folder_size_dicts, csv_writer):

    print "  Writing storage usage data"

    for row_dict in folder_size_dicts:
        csv_writer.writerow(row_dict)


#=====
#
# SCRIPT BODY
#
#=====

parser = argparse.ArgumentParser()

parser.add_argument("billing_config_file",
                    help='The BillingConfig file')
parser.add_argument("-r", "--billing_root",
                    default=None,
                    help='The Billing Root directory [default = None]')
parser.add_argument("--no_usage", action="store_true",
                    default=False,
                    help="Don't run storage usage calculations [default = false]")
parser.add_argument("-s", "--storage_usage_csv",
                    default=None,
                    help="The storage usage CSV file.")
parser.add_argument("-y", "--year", type=int, choices=range(2013, 2031),
                    default=None,
                    help="The year to be filtered out. [default = this year]")
parser.add_argument("-m", "--month", type=int, choices=range(1, 13),
                    default=None,
                    help="The month to be filtered out. [default = last month]")
parser.add_argument("-v", "--verbose", action="store_true",
                    default=False,
                    help='Get chatty [default = false]')
parser.add_argument("-d", "--debug", action="store_true",
                    default=False,
                    help='Get REAL chatty [default = false]')

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
if not os.path.exists(year_month_dir):
    os.makedirs(year_month_dir)

# Generate storage usage pathname.
if args.storage_usage_csv is not None:
    storage_usage_pathname = args.storage_usage_csv
else:
    storage_usage_filename = "%s.%s-%02d.csv" % (STORAGE_PREFIX, str(year), month)
    storage_usage_pathname = os.path.join(year_month_dir, storage_usage_filename)

#
# Generate storage usage data.
#
folder_size_dicts = compute_storage_charges(billing_config_wkbk, begin_month_timestamp, end_month_timestamp)

#
# Output storage usage data into a CSV.
#
output_storage_usage_csv_file = open(storage_usage_pathname, 'w')
csv_writer = csv.DictWriter(output_storage_usage_csv_file, BILLING_DETAILS_SHEET_COLUMNS['Storage'])

csv_writer.writeheader()
write_storage_usage_data(folder_size_dicts,csv_writer)