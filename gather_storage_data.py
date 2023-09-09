#!/usr/bin/env python3

#===============================================================================
#
# gather_storage_data.py - Gather the usage and quota information all the folders in the
#                           BillingConfig file and output them into a CSV file.
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
#     rows of the form Date, Timestamp, Folder, Size, Used, Inode Quota, Inodes Used.
#
# ASSUMPTIONS:
#   Script is run on a machine which can run the quota command, if needed.
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
import re
import sys
import subprocess

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
global QUOTA_EXECUTABLE
global USAGE_EXECUTABLE
global INODES_EXECUTABLE
global STORAGE_BLOCK_SIZE_ARG
global STORAGE_PREFIX
global BILLING_DETAILS_SHEET_COLUMNS
global TOPLEVEL_DIRECTORIES
global BAAS_SUBDIR_NAME

#=====
#
# GLOBAL VARS
#
#=====
global folder_storage_data_dict

#
# FUNCTIONS
#
#=====
# In billing_common.py
global sheet_get_named_column
global from_timestamp_to_excel_date
global from_excel_date_to_timestamp
global from_ymd_date_to_timestamp
global from_datetime_to_timestamp
global read_config_sheet
global argparse_get_parent_parser
global argparse_get_year_month
global argparse_get_billingroot_billingconfig

# Look for a date in the filename for a storage data file.
# Expected substrings in filename:
# YYYY.MM.DD
# If found, return timestamp for that date.
# If not, return timestamp for now.
#
def parse_filename_for_timestamp(filename):

    # Only look in basename of filename for date
    basename = os.path.basename(filename)

    match = re.search("(\d{2,4})\.(\d{2})\.(\d{2})", basename)
    if match is not None:
        year = int(match.group(1))
        if len(match.group(1)) == 2:
            year += 2000
        month = int(match.group(2))
        day = int(match.group(3))

        return datetime.datetime(year, month, day).timestamp()
    else:
        return datetime.datetime.now().timestamp()


# Reads in a storage data file (.csv)
# Format: folder_basename,used_gb,quota_gb,inodes_used,inodes_quota
# Stores the last four values in a dictionary indexed by the first value
def read_storage_data_file(storage_data_filename):

    # Parse filename for possible date held within.
    timestamp_for_data = parse_filename_for_timestamp(storage_data_filename)

    with open(storage_data_filename) as storage_fp:
        # Read header line
        header_line = storage_fp.readline()

        # Parse header into fields
        header_list = header_line.split(',')

        if len(header_list) != 5:
            print(file=sys.stderr)
            print("  Header", header_line, "does not have five columns", file=sys.stderr)
            return False

        if header_list[0] == "project":
            folder_prefix = "/projects/"
        elif header_list[0] == "PI":
            folder_prefix = "/labs/"
        else:
            folder_prefix = ""

        for data_line in storage_fp:
            data_line = data_line[:-1]  # remove trailing newline
            (folder,used_gb,quota_gb,inodes_used,inodes_quota) = data_line.split(',')

            used_tb = int(used_gb)/(1024**3)
            quota_tb = int(quota_gb)/(1024**3)

            inodes_used = int(inodes_used)
            inodes_quota = int(inodes_quota)

            # Store folder storage data
            folder_storage_data_dict[folder_prefix + folder] = (timestamp_for_data,used_tb,quota_tb,inodes_used,inodes_quota)

    return True

# Gets the quota for the given PI tag.
# Returns a tuple of (size used, quota) with values in Tb, or
# None if there was a problem parsing the quota command output.
def get_folder_quota(machine, folder):

    # Build and execute the quota command.
    quota_cmd = []

    # Add ssh if this is a remote call.
    if machine is not None:
        quota_cmd += ["ssh", machine]

    # Build the quota command from the assembled arguments.
    quota_cmd += QUOTA_EXECUTABLE + STORAGE_BLOCK_SIZE_ARG + [folder]

    try:
        quota_output = subprocess.check_output(quota_cmd, text=True, encoding="utf-8")
    except subprocess.CalledProcessError as cpe:
        print("Couldn't get quota for %s (exit %d)" % (folder, cpe.returncode), file=sys.stderr)
        print(" Command:", quota_cmd, file=sys.stderr)
        print(" Output:", cpe.output, file=sys.stderr)
        return None

    # Parse the results.
    for line in quota_output.split('\n'):

        fields = line.split()
        if fields[2] != "Used":
            used  = int(fields[2])
            quota = int(fields[1])

            return (used/1024.0, quota/1024.0)  # Return values in fractions of Tb.
    else:
        return None


# Gets the usage for the given folder.
# Returns a tuple of (quota, quota) with values in Tb, or
# None if there was a problem parsing the usage command output.
def get_folder_usage(machine, folder):

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
        print(" Output:", cpe.output, file=sys.stderr)
        return None

    # Parse the results.
    for line in usage_output.split('\n'):

        fields = line.split()
        used = int(fields[0])

        return (used/1024.0, used/1024.0)  # Return values in fractions of Tb.

    return None


def get_folder_inodes(machine, folder):

    # Build and execute the inodes command.
    inodes_cmd = []

    # Add ssh if this is a remote call.
    if machine is not None:
        inodes_cmd += ["ssh", machine]

    inodes_cmd += INODES_EXECUTABLE + [folder]

    try:
        inodes_output = subprocess.check_output(inodes_cmd, text=True, encoding="utf-8")
    except subprocess.CalledProcessError as cpe:
        print("Couldn't get inodes for %s (exit %d)" % (folder, cpe.returncode), file=sys.stderr)
        print(" Command:", inodes_cmd, file=sys.stderr)
        print(" Output:", cpe.output, file=sys.stderr)
        return None

    # Parse the results.
    for line in inodes_output.split('\n'):

        fields = line.split()
        if fields[1] != "Inodes":  # the header line
            inodes_used = int(fields[2])
            inodes_quota = int(fields[3])

            return inodes_used, inodes_quota

    return None


# Generates the Storage sheet data from folder quotas and usages.
# Returns mapping from folders to [timestamp, total, used]
def compute_storage_charges(config_wkbk, begin_timestamp, end_timestamp):

    print()
    print("COMPUTING STORAGE CHARGES...")

    # Lists of folders to measure come from:
    #  "PI Folder" column of "PIs" sheet, and
    #  "Folder" column of "Folders" sheet.

    # Get lists of folders, quota booleans from PIs sheet.
    #pis_sheet = config_wkbk.sheet_by_name('PIs')
    pis_sheet = config_wkbk['PIs']

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
    # folder_sheet = config_wkbk.sheet_by_name('Folders')
    folder_sheet = config_wkbk['Folders']

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

    folder_aggregate_rows = list(zip(folders, pi_tags, measure_types, dates_added, dates_remvd))
    sorted_folder_aggregate_rows = sorted(folder_aggregate_rows, key = lambda x: x[0] if x[0] is not None else '')

    # Create mapping from folders to space used.
    for (folder, pi_tag, measure_type, date_added, date_removed) in sorted_folder_aggregate_rows:

        # Skip measuring this folder entry if the folder is "None".
        if folder is None or folder.startswith('None'): continue

        # Account for multiple folders separated by commas.
        folder_list = folder.split(',')

        for this_folder in folder_list:

            # Skip measuring this folder if we have already done it.
            if this_folder in folders_measured:
                continue

            # If this folder has been added prior to or within this month
            # and has not been removed before the beginning of this month, analyze it.
            if (end_timestamp > from_datetime_to_timestamp(date_added) and
                (date_removed == '' or date_removed is None or begin_timestamp < from_datetime_to_timestamp(date_removed)) ):

                # Split folder into machine:dir components.
                if ':' in this_folder:
                    (this_folder_machine, this_folder_dir) = this_folder.split(':')
                else:
                    this_folder_machine = None
                    this_folder_dir = this_folder

                #
                # Get storage data for current folder
                #
                # If it is already read in (folder_storage_data_dict), then get it from there
                # If not and measure_type is "quota", generate a new entry in folder_storage_data_dict for the quota of the folder
                # If not and measure_type is "usage",  generate a new entry in folder_storage_data_dict for the usage of the folder
                # If not, then mention we have no data for folder.
                #
                if this_folder in folder_storage_data_dict:
                    print(this_folder, ": Found in database")

                elif this_folder.lower() in folder_storage_data_dict:
                    print(this_folder, ": Found in database")

                    folder_storage_data_dict[this_folder] = folder_storage_data_dict[this_folder.lower()]

                elif measure_type == "quota":
                    # Check folder's quota.
                    print(this_folder, ": Getting quota", end=' ')
                    quota_tuple = get_folder_quota(this_folder_machine, this_folder_dir)
                    if quota_tuple is None:
                        print("Could not get quota for %s...SKIPPING measurement" % this_folder_dir, file=sys.stderr)
                        continue

                    # Check folder's inodes
                    print("inodes")  # end status output line
                    inodes_tuple = get_folder_inodes(this_folder_machine, this_folder_dir)
                    if inodes_tuple is None:
                        print("Could not get inodes for %s...SKIPPING measurement" % this_folder_dir, file=sys.stderr)
                        continue

                    # Set entry variables for an addition to the database
                    folder_timestamp = datetime.datetime.now().timestamp()
                    (used_tb, quota_tb) = quota_tuple
                    (inodes_used, inodes_quota) = inodes_tuple

                    folder_storage_data_dict[this_folder] = (folder_timestamp, used_tb, quota_tb, inodes_used, inodes_quota)

                elif measure_type == "usage" and not args.no_usage:
                    # Check folder's usage.
                    print(this_folder, "Measuring usage")

                    usage_tuple = get_folder_usage(this_folder_machine, this_folder_dir)
                    if usage_tuple is None:
                        print("Could not get usage for %s...SKIPPING measurement" % this_folder_dir, file=sys.stderr)

                    # Set entry variables for an addition to the database
                    folder_timestamp = datetime.datetime.now().timestamp()
                    (used_tb, quota_tb) = usage_tuple

                    folder_storage_data_dict[this_folder] = (folder_timestamp, used_tb, quota_tb, 0, 0)

                else:
                    # Use null values for no usage data.
                    print("SKIPPING measurement for", this_folder)
                    continue


                (folder_timestamp, used_tb, quota_tb, inodes_used, inodes_quota) = folder_storage_data_dict[this_folder]

                folder_size_dicts.append({ 'Date Measured' : from_timestamp_to_excel_date(folder_timestamp),
                                           'Timestamp' : folder_timestamp,
                                           'Folder' : this_folder,
                                           'Size' : quota_tb,
                                           'Used' : used_tb,
                                           'Inodes Quota': inodes_quota,
                                           'Inodes Used': inodes_used
                                           })

                # Add the folder measured to a set of folders we have measured.
                folders_measured.add(this_folder)

            else:
                print("  *** Excluding %s for PI %s: folder not active in this month" % (folder, pi_tag))

    return folder_size_dicts


# Write storage usage data into a storage usage CSV file.
# Takes in mapping from folders to [timestamp, total, used].
def write_storage_usage_data(folder_size_dicts, csv_writer):

    print("WRITING STORAGE USAGE DATA...")
    for row_dict in folder_size_dicts:
        csv_writer.writerow(row_dict)


#=====
#
# SCRIPT BODY
#
#=====

parser = argparse.ArgumentParser(parents=[argparse_get_parent_parser()])

parser.add_argument("storage_data_files", nargs="*",
                    default=[],
                    help="Storage data files [optional]")
parser.add_argument("--no_usage", action="store_true",
                    default=False,
                    help="Don't run storage usage calculations [default = false]")
parser.add_argument("--include_baas_folders", action="store_true",
                    default=False,
                    help="In PI Folders, also measure 'FOLDER/BaaS' [default = false]")
parser.add_argument("-s", "--storage_usage_csv",
                    default=None,
                    help="The storage usage CSV file.")

args = parser.parse_args()

#
# Process arguments.
#

# Get year/month-related arguments
(year, month, begin_month_timestamp, end_month_timestamp) = argparse_get_year_month(args)

# Get BillingRoot and BillingConfig arguments
(billing_root, billing_config_file) = argparse_get_billingroot_billingconfig(args, year, month)

#
# Open the Billing Config workbook.
#
# billing_config_wkbk = xlrd.open_workbook(billing_config_file)
billing_config_wkbk = openpyxl.load_workbook(billing_config_file)

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
print("GATHERING STORAGE DATA FOR %02d/%d:" % (month, year))
print("  BillingConfigFile: %s" % billing_config_file)
print("  BillingRoot: %s" % billing_root)
print()
if args.no_usage:
    print("  Not measuring storage usage figures")
    print()
print("  Storage usage file to be output: %s" % storage_usage_pathname)
print()

#
# Read in any storage data files.
#
folder_storage_data_dict = dict()   # Storage data file data will be stored here

print()
for storage_file in args.storage_data_files:
    print("READING STORAGE DATA FILE", storage_file)
    read_storage_data_file(storage_file)

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

print()
print("STORAGE DATA GATHERING COMPLETE.")
