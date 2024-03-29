#!/usr/bin/env python3

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
global QUOTA_EXECUTABLE_GPFS
global QUOTA_EXECUTABLE_ISILON
global USAGE_EXECUTABLE
global STORAGE_BLOCK_SIZE_ARG
global STORAGE_PREFIX
global BILLING_DETAILS_SHEET_COLUMNS
global GPFS_TOPLEVEL_DIRECTORIES
global ISILON_TOPLEVEL_DIRECTORIES
global BAAS_SUBDIR_NAME

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
global argparse_get_parent_parser
global argparse_get_year_month
global argparse_get_billingroot_billingconfig

# Deduces the fileset that a folder being measured by quota lives in.
#
# Patterns:
#  /srv/gsfs0/DIR1      - Device "gsfs0", Fileset "DIR1"
#  /srv/gsfs0/DIR1/DIR2 - Device "gsfs0", Fileset "DIR1.DIR2"
#  /srv/gsfs0/BaaS/Labs/DIR1 - Device "gsfs0", Fileset "BaaS.DIR1"
#
# Default device: "ifs"
#
def get_device_and_fileset_from_folder(folder):

    # We have a list of top-level directories for the GPFS and Isilon systems, respectively.
    if 0 in [folder.find(d) for d in GPFS_TOPLEVEL_DIRECTORIES]:

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
            print("get_device_and_fileset_from_folder(): Path %s is not long enough" % folder, file=sys.stderr)
            return None

    elif 0 in [folder.find(d) for d in ISILON_TOPLEVEL_DIRECTORIES]:

        device = "ifs"
        fileset = folder

        return (device, fileset)

    else:

        print("get_device_and_fileset_from_folder(): Cannot get device and fileset from %s" % folder, file=sys.stderr)
        return None


def get_folder_quota_from_gpfs(machine, device, fileset):

    # Build and execute the quota command.
    quota_cmd = []

    # Add ssh if this is a remote call.
    if machine is not None:
        quota_cmd += ["ssh", machine]

    # Build the quota command from the assembled arguments.
    quota_cmd += QUOTA_EXECUTABLE_GPFS + [fileset] + STORAGE_BLOCK_SIZE_ARG + [device]

    try:
        quota_output = subprocess.check_output(quota_cmd, text=True, encoding="utf-8")
    except subprocess.CalledProcessError as cpe:
        print("Couldn't get quota for %s (exit %d)" % (fileset, cpe.returncode), file=sys.stderr)
        print(" Command:", quota_cmd, file=sys.stderr)
        print(" Output: %s" % (cpe.output), file=sys.stderr)
        return None

    # Parse the results.
    for line in quota_output.split('\n'):

        fields = line.split()

        # If the first word on this line is 'gsfs0', this is the line we want.
        if fields[0] == 'gsfs0':
            used  = int(fields[2])
            quota = int(fields[4])

            return (used/1024.0, quota/1024.0)  # Return values in fractions of Tb.
    else:
        return None


def get_folder_quota_from_isilon(machine, device, fileset):

    # Build and execute the quota command.
    quota_cmd = []

    # Add ssh if this is a remote call.
    if machine is not None:
        quota_cmd += ["ssh", machine]

    # Build the quota command from the assembled arguments.
    quota_cmd += QUOTA_EXECUTABLE_ISILON + STORAGE_BLOCK_SIZE_ARG + [fileset]

    try:
        quota_output = subprocess.check_output(quota_cmd, text=True, encoding="utf-8")
    except subprocess.CalledProcessError as cpe:
        print("Couldn't get quota for %s (exit %d)" % (fileset, cpe.returncode), file=sys.stderr)
        print(" Command:", quota_cmd, file=sys.stderr)
        print(" Output: %s" % (cpe.output), file=sys.stderr)
        return None

    # Parse the results.
    for line in quota_output.split('\n'):

        fields = line.split()

        # If the first word on this line is not "Filesystem", this is the line we want.
        if fields[0].find("Filesystem") == -1:
            used  = int(fields[2])
            quota = int(fields[1])

            return (used/1024.0, quota/1024.0)  # Return values in fractions of Tb.
    else:
        return None


# Gets the quota for the given PI tag.
# Returns a tuple of (size used, quota) with values in Tb, or
# None if there was a problem parsing the quota command output.
def get_folder_quota(machine, folder, pi_tag):

    if args.verbose: print("  Getting folder quota for %s..." % (pi_tag))

    # Find the fileset to get the quota of, from the folder name.
    device_and_fileset = get_device_and_fileset_from_folder(folder)
    if device_and_fileset is None:
        print("ERROR: No fileset for folder %s; ignoring..." % (folder), file=sys.stderr)
        return None

    (device, fileset) = device_and_fileset

    if device == "gsfs0":
        return get_folder_quota_from_gpfs(machine, device, fileset)
    elif device == "ifs":
        return get_folder_quota_from_isilon(machine, device, fileset)
    else:
        return None


# Gets the usage for the given folder.
# Returns a tuple of (quota, quota) with values in Tb, or
# None if there was a problem parsing the usage command output.
def get_folder_usage(machine, folder, pi_tag):

    if args.verbose: print("  Getting folder usage of %s..." % (folder))

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
        usage_output = subprocess.check_output(usage_cmd, text=True, encoding="utf-8")
    except subprocess.CalledProcessError as cpe:
        print("Couldn't get usage for %s (exit %d)" % (folder, cpe.returncode), file=sys.stderr)
        print(" Command:", usage_cmd, file=sys.stderr)
        print(" Output: %s" % (cpe.output), file=sys.stderr)
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

    print("Computing storage charges...")

    # Lists of folders to measure come from:
    #  "PI Folder" column of "PIs" sheet, and
    #  "Folder" column of "Folders" sheet.

    # Get lists of folders, quota booleans from PIs sheet.
    pis_sheet = config_wkbk.sheet_by_name('PIs')

    pis_sheet_folders     = sheet_get_named_column(pis_sheet, 'PI Folder')
    pis_sheet_pi_tags     = sheet_get_named_column(pis_sheet, 'PI Tag')
    pis_sheet_measure_types = ['quota'] * len(pis_sheet_folders)   # All PI folders are measured by quota.
    pis_sheet_dates_added = sheet_get_named_column(pis_sheet, 'Date Added')
    pis_sheet_dates_remvd = sheet_get_named_column(pis_sheet, 'Date Removed')

    # Potentially add "BaaS" subfolders to PI folder names, if switch given.
    if args.include_baas_folders:
        pis_sheet_baas_folders = [x + "/" + BAAS_SUBDIR_NAME for x in pis_sheet_folders]
        pis_sheet_baas_pi_tags = pis_sheet_pi_tags
        pis_sheet_baas_measure_types = ['usage'] * len(pis_sheet_baas_folders)   # All PI BaaS folders are measured by usage.
        pis_sheet_baas_dates_added = pis_sheet_dates_added
        pis_sheet_baas_dates_remvd = pis_sheet_dates_remvd
    else:
        pis_sheet_baas_folders = []
        pis_sheet_baas_pi_tags = []
        pis_sheet_baas_measure_types = []
        pis_sheet_baas_dates_added = []
        pis_sheet_baas_dates_remvd = []

    # Get lists of folders, quota booleans from Folders sheet.
    folder_sheet = config_wkbk.sheet_by_name('Folders')

    folders_sheet_folders     = sheet_get_named_column(folder_sheet, 'Folder')
    folders_sheet_pi_tags     = sheet_get_named_column(folder_sheet, 'PI Tag')
    folders_sheet_measure_types = sheet_get_named_column(folder_sheet, 'Method')
    folders_sheet_dates_added = sheet_get_named_column(folder_sheet, 'Date Added')
    folders_sheet_dates_remvd = sheet_get_named_column(folder_sheet, 'Date Removed')

    # Assemble the lists from above.
    folders       = pis_sheet_folders + pis_sheet_baas_folders + folders_sheet_folders
    pi_tags       = pis_sheet_pi_tags + pis_sheet_baas_pi_tags + folders_sheet_pi_tags
    measure_types = pis_sheet_measure_types + pis_sheet_baas_measure_types + folders_sheet_measure_types
    dates_added   = pis_sheet_dates_added + pis_sheet_baas_dates_added + folders_sheet_dates_added
    dates_remvd   = pis_sheet_dates_remvd + pis_sheet_baas_dates_remvd + folders_sheet_dates_remvd

    # List of dictionaries with keys from BILLING_DETAILS_SHEET_COLUMNS['Storage'].
    folder_size_dicts = []
    # Set of folders that have been measured
    folders_measured = set()

    # Create mapping from folders to space used.
    for (folder, pi_tag, measure_type, date_added, date_removed) in zip(folders, pi_tags, measure_types, dates_added, dates_remvd):

        # Skip measuring this folder entry if the folder is None.
        if folder.startswith('None'): continue

        # Account for multiple folders separated by commas.
        pi_folder_list = folder.split(',')

        for pi_folder in pi_folder_list:

            # Skip measuring this folder if we have already done it.
            if pi_folder in folders_measured: continue

            # If this folder has been added prior to or within this month
            # and has not been removed before the beginning of this month, analyze it.
            if (end_timestamp > from_excel_date_to_timestamp(date_added) and
                (date_removed == '' or begin_timestamp < from_excel_date_to_timestamp(date_removed)) ):

                # Split folder into machine:dir components.
                if ':' in pi_folder:
                    (machine, dir) = pi_folder.split(':')
                else:
                    machine = None
                    dir = pi_folder

                if measure_type == "quota":
                    # Check folder's quota.
                    print("Getting quota for %s" % pi_folder)

                    used_and_total = get_folder_quota(machine, dir, pi_tag)
                    if used_and_total is None:
                        print("Could not get %s for %s...SKIPPING measurement" % (measure_type, dir), file=sys.stderr)

                elif measure_type == "usage" and not args.no_usage:
                    # Check folder's usage.
                    print("Measuring usage for %s" % pi_folder)

                    used_and_total = get_folder_usage(machine, dir, pi_tag)
                    if used_and_total is None:
                        print("Could not get %s for %s...SKIPPING measurement" % (measure_type, dir), file=sys.stderr)

                else:
                    # Use null values for no usage data.
                    print("SKIPPING measurement for %s" % pi_folder)
                    used_and_total = None

                if used_and_total is None:
                    used = total = 0
                else:
                    (used, total) = used_and_total

                # Record when we measured the storage.
                measured_timestamp = time.time()

                folder_size_dicts.append({ 'Date Measured' : from_timestamp_to_excel_date(measured_timestamp),
                                           'Timestamp' : measured_timestamp,
                                           'Folder' : pi_folder,
                                           'Size' : total,
                                           'Used' :used })

                # Add the folder measured to a set of folders we have measured.
                folders_measured.add(pi_folder)

            else:
                print("  *** Excluding %s for PI %s: folder not active in this month" % (folder, pi_tag))

    return folder_size_dicts


# Write storage usage data into a storage usage CSV file.
# Takes in mapping from folders to [timestamp, total, used].
def write_storage_usage_data(folder_size_dicts, csv_writer):

    print("  Writing storage usage data")

    for row_dict in folder_size_dicts:
        csv_writer.writerow(row_dict)


#=====
#
# SCRIPT BODY
#
#=====

parser = argparse.ArgumentParser(parents=[argparse_get_parent_parser()])

parser.add_argument("billing_config_file",
                    help='The BillingConfig file')
parser.add_argument("-r", "--billing_root",
                    default=None,
                    help='The Billing Root directory [default = None]')
parser.add_argument("--no_usage", action="store_true",
                    default=False,
                    help="Don't run storage usage calculations [default = false]")
parser.add_argument("--include_baas_folders", action="store_true",
                    default=False,
                    help="In PI Folders, also measure 'FOLDER/BaaS' [default = false]")
parser.add_argument("-s", "--storage_usage_csv",
                    default=None,
                    help="The storage usage CSV file.")
parser.add_argument("-y", "--year", type=int, choices=list(range(2013, 2031)),
                    default=None,
                    help="The year to be filtered out. [default = this year]")
parser.add_argument("-m", "--month", type=int, choices=list(range(1, 13)),
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

# Get year/month-related arguments
(year, month, begin_month_timestamp, end_month_timestamp) = argparse_get_year_month(args)

# Get BillingRoot and BillingConfig arguments
(billing_root, billing_config_file) = argparse_get_billingroot_billingconfig(args)

# Open the BillingConfig workbook
billing_config_wkbk = openpyxl.load_workbook(billing_config_file)  # , read_only=True)

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
# Output the state of arguments.
#
print("SNAPSHOTTING STORAGE FOR %02d/%d:" % (month, year))
print("  BillingConfigFile: %s" % billing_config_file)
print("  BillingRoot: %s" % billing_root)
print()
if args.no_usage:
    print("  Not recording storage usage figures")
    print()
print("  Storage usage file to be output: %s" % storage_usage_pathname)
print()

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
