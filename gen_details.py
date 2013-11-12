#!/usr/bin/env python

#===============================================================================
#
# gen_details.py - Generate billing details for month/year.
#
# ARGS:
#   1st: the BillingConfig spreadsheet.
#
# SWITCHES:
#   --accounting_file: Location of accounting file (overrides BillingConfig.xlsx)
#   --billing_root:    Location of BillingRoot directory (overrides BillingConfig.xlsx)
#                      [default if no BillingRoot in BillingConfig.xlsx or switch given: CWD]
#   --year:            Year of snapshot requested. [Default is this year]
#   --month:           Month of snapshot requested. [Default is last month]
#   --no_storage:      Don't run the storage calculations.
#   --no_usage:        Don't run the storage usage calculations (only the quotas).
#   --no_computing:    Don't run the computing calculations.
#   --no_consulting:   Don't run the consulting calculations.
#   --scg3:            Add 'scg3' to list of hostname prefixes for billable jobs.
#
# INPUT:
#   BillingConfig spreadsheet.
#   SGE Accounting snapshot file (from snapshot_accounting.py).
#     - Expected in BillingRoot/<year>/<month>/SGEAccounting.<year>-<month>.xlsx
#
# OUTPUT:
#   BillingDetails spreadsheet in BillingRoot/<year>/<month>/BillingDetails.<year>-<month>.xlsx
#   Various messages about current processing status to STDOUT.
#
# ASSUMPTIONS:
#   Depends on xlrd and xlsxwriter modules.
#   The input spreadsheet has been certified by check_config.py.
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
import time

import xlrd
import xlsxwriter

# Simulate an "include billing_common.py".
SCRIPT_DIR = os.path.dirname(os.path.realpath(__file__))
execfile(os.path.join(SCRIPT_DIR, "billing_common.py"))

#=====
#
# CONSTANTS
#
#=====


#=====
#
# GLOBALS
#
#=====
# In billing_common.py
global SGEACCOUNTING_PREFIX
global BILLING_DETAILS_PREFIX

# Commands for determining folder quotas and usages.
QUOTA_EXECUTABLE = ['/usr/lpp/mmfs/bin/mmlsquota', '-j']
USAGE_EXECUTABLE = ['sudo', 'du', '-s']
STORAGE_BLOCK_SIZE_ARG = ['--block-size=1G']  # Works in both above commands.

#
# Formats for the output workbook, to be initialized along with the workbook.
#
BOLD_FORMAT = None
DATE_FORMAT = None
INT_FORMAT  = None
FLOAT_FORMAT = None
MONEY_FORMAT = None
PERCENT_FORMAT = None

#=====
#
# FUNCTIONS
#
#=====

# In billing_common.py
global read_config_sheet
global sheet_get_named_column
global sheet_name_to_sheet
global from_timestamp_to_excel_date
global from_excel_date_to_timestamp
global from_timestamp_to_date_string
global from_excel_date_to_date_string
global from_ymd_date_to_timestamp

# Initialize the output BillingDetails workbook, given as argument.
# It creates all the formats used within the workbook, and saves them
# as the global variables listed at the top of the method.
# It also creates all the sheets within the workbook, and the column
# headers within those sheets.  The method returns a dict of mappings
# from sheet_name to the workbook's Sheet object for that name.
def init_billing_details_wkbk(workbook):

    global BOLD_FORMAT
    global DATE_FORMAT
    global FLOAT_FORMAT
    global INT_FORMAT
    global MONEY_FORMAT
    global PERCENT_FORMAT

    global BILLING_DETAILS_SHEET_COLUMNS  # In billing_common.py

    # Create formats for use within the workbook.
    BOLD_FORMAT = workbook.add_format({'bold' : True})
    DATE_FORMAT = workbook.add_format({'num_format' : 'mm/dd/yy'})
    INT_FORMAT  = workbook.add_format({'num_format' : '0'})
    FLOAT_FORMAT = workbook.add_format({'num_format' : '0.0'})
    MONEY_FORMAT = workbook.add_format({'num_format' : '$0.00'})
    PERCENT_FORMAT = workbook.add_format({'num_format' : '0%'})

    sheet_name_to_sheet = dict()

    for sheet_name in BILLING_DETAILS_SHEET_COLUMNS:

        sheet = workbook.add_worksheet(sheet_name)
        for col in range(0, len(BILLING_DETAILS_SHEET_COLUMNS[sheet_name])):
            sheet.write(0, col, BILLING_DETAILS_SHEET_COLUMNS[sheet_name][col], BOLD_FORMAT)

        sheet_name_to_sheet[sheet_name] = sheet

    # Make the Storage sheet the active one.
    sheet_name_to_sheet['Storage'].activate()

    return sheet_name_to_sheet

# Gets the quota for the given PI tag.
# Returns a tuple of (size used, quota) with values in Tb, or
# None if there was a problem parsing the quota command output.
def get_folder_quota(folder, pi_tag):

    print "  Getting folder quota for %s..." % (pi_tag)

    # Build and execute the quota command.
    quota_cmd = QUOTA_EXECUTABLE + ["projects." + pi_tag] + STORAGE_BLOCK_SIZE_ARG + ["gsfs0"]

    try:
        quota_output = subprocess.check_output(quota_cmd)
    except subprocess.CalledProcessError as cpe:
        print >> sys.stderr, "Couldn't get quota for %s (exit %d)" % (pi_tag, cpe.returncode)
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
def get_folder_usage(folder, pi_tag):

    print "  Getting folder usage of %s..." % (folder)

    # Build and execute the quota command.
    usage_cmd = USAGE_EXECUTABLE + STORAGE_BLOCK_SIZE_ARG + [folder]

    try:
        usage_output = subprocess.check_output(usage_cmd)
    except subprocess.CalledProcessError as cpe:
        print >> sys.stderr, "Couldn't get usage for %s (exit %d)" % (folder, cpe.returncode)
        print >> sys.stderr, " Output: %s" % (cpe.output)
        return None

    # Parse the results.
    for line in usage_output.split('\n'):

        fields = line.split()
        used = int(fields[0])

        return (used/1024.0, used/1024.0)  # Return values in fractions of Tb.

    return None

# Given a list of tuples of job details, writes the job details to
# the sheet given.  Possible sheets for use in this method are
# typically the "Computing", "Nonbillable Jobs", and "Failed Jobs" sheets.
def write_job_details(sheet, job_details):

    for row in range(0, len(job_details)):

        # Bump rows down below header line.
        sheet_row = row + 1

        # A little feedback for the people.
        if not args.verbose:
            if sheet_row % 1000 == 0:
                sys.stdout.write('.')
                sys.stdout.flush()

        # 'Job Date'
        col = 0
        sheet.write(sheet_row, col, from_timestamp_to_excel_date(job_details[row][col]), DATE_FORMAT)

        # 'Job Timestamp'
        col += 1
        sheet.write(sheet_row, col, job_details[row][col], INT_FORMAT)
        if args.verbose: print job_details[row][col],

        # 'Username'
        col += 1
        sheet.write(sheet_row, col, job_details[row][col])
        if args.verbose: print job_details[row][col],

        # 'Job Name'
        col += 1
        sheet.write(sheet_row, col, job_details[row][col])
        if args.verbose: print job_details[row][col],

        # 'Account'
        col += 1
        sheet.write(sheet_row, col, job_details[row][col])
        if args.verbose: print job_details[row][col],

        # 'Node'
        col += 1
        sheet.write(sheet_row, col, job_details[row][col])
        if args.verbose: print job_details[row][col],

        # 'Slots'
        col += 1
        sheet.write(sheet_row, col, job_details[row][col], INT_FORMAT)
        if args.verbose: print job_details[row][col],

        # 'Wallclock'
        col += 1
        sheet.write(sheet_row, col, job_details[row][col], INT_FORMAT)
        if args.verbose: print job_details[row][col],

        # 'JobID'
        col += 1
        sheet.write(sheet_row, col, job_details[row][col], INT_FORMAT)
        if args.verbose: print job_details[row][col],

        # Extra column if needed: 'Reason' or 'Failed Code'
        if col < len(job_details[row])-1:
            col += 1
            sheet.write(sheet_row, col, job_details[row][col])
            if args.verbose: print job_details[row][col],

        if args.verbose: print

    print

# Generates the Storage sheet data from folder quotas and usages.
def compute_storage_charges(config_wkbk, begin_timestamp, end_timestamp, storage_sheet):

    print "Computing storage charges..."

    # Get lists of folders, quota booleans.
    folder_sheet = config_wkbk.sheet_by_name('Folders')

    folders     = sheet_get_named_column(folder_sheet, 'Folder')
    pi_tags     = sheet_get_named_column(folder_sheet, 'PI Tag')
    quota_bools = sheet_get_named_column(folder_sheet, 'By Quota?')
    dates_added = sheet_get_named_column(folder_sheet, 'Date Added')

    folder_sizes = []

    # Create mapping from folders to space used.
    for (folder, pi_tag, quota_bool, date_added) in zip(folders, pi_tags, quota_bools, dates_added):

        # If this folder has been added prior to or within this month, analyze it.
        if end_timestamp > from_excel_date_to_timestamp(date_added):
            if quota_bool == "yes":
                # Check folder's quota.
                (used, total) = get_folder_quota(folder, pi_tag)
            else:
                if not args.no_usage:
                    # Check folder's usage.
                    (used, total) = get_folder_usage(folder, pi_tag)
                else:
                    # Use null values for no usage data.
                    (used, total) = (0, 0)

            folder_sizes.append([ time.time(), folder, total, used ])
        else:
            print "  *** Excluding %s for PI %s: not active in this month" % (folder, pi_tag)

    # Write space-used mapping into details workbook.
    for row in range(0, len(folder_sizes)):

        sheet_row = row + 1

        # 'Date Measured'
        col = 0
        storage_sheet.write(sheet_row, col, from_timestamp_to_excel_date(folder_sizes[row][col]), DATE_FORMAT)

        # 'Timestamp'
        col += 1
        storage_sheet.write(sheet_row, col, folder_sizes[row][col-1])
        if args.verbose: print folder_sizes[row][col-1],

        # 'Folder'
        col += 1
        storage_sheet.write(sheet_row, col, folder_sizes[row][col-1])
        if args.verbose: print folder_sizes[row][col-1],

        # 'Size'
        col += 1
        storage_sheet.write(sheet_row, col, folder_sizes[row][col-1], FLOAT_FORMAT)
        if args.verbose: print folder_sizes[row][col-1],

        # 'Used'
        col += 1
        storage_sheet.write(sheet_row, col, folder_sizes[row][col-1], FLOAT_FORMAT)
        if args.verbose: print folder_sizes[row][col-1]


# Generates the job details stored in the "Computing", "Nonbillable Jobs", and "Failed Jobs" sheets.
def compute_computing_charges(config_wkbk, begin_timestamp, end_timestamp, accounting_file,
                              computing_sheet, nonbillable_job_sheet, failed_job_sheet):

    # In billing_common.py
    global ACCOUNTING_FIELDS
    global ACCOUNTING_FAILED_CODES
    global BILLABLE_HOSTNAME_PREFIXES

    print "Computing computing charges..."

    # Read in the Usernames from the Users sheet.
    users_sheet = config_wkbk.sheet_by_name('Users')
    users_list = sheet_get_named_column(users_sheet, "Username")
    #  NOTE: This column may have some duplicates in it.
    #        Need to make a set out of the result.
    users_list = set(users_list)

    # Read in the PI Tag list from the PIs sheet.
    pis_sheet = config_wkbk.sheet_by_name('PIs')
    pi_tag_list = sheet_get_named_column(pis_sheet, 'PI Tag')

    # Read in the Job Tags from the Job Tags sheet.
    job_tags_sheet = config_wkbk.sheet_by_name('JobTags')
    job_tag_list = sheet_get_named_column(job_tags_sheet, "Job Tag")

    print "  Reading accounting file %s" % (os.path.abspath(accounting_file))

    #
    # Open the current accounting file for input.
    #
    accounting_fp = open(accounting_file, "r")

    #
    # Read all the lines of the current accounting file.
    #  Output to the details spreadsheet those jobs
    #  which have "end_times" in the given month,
    #  and "owner"s in the list of users.
    #
    not_in_users_list = set()
    not_in_job_tag_list = set()

    failed_job_details           = []  # Jobs which failed.
    billable_job_details         = []  # Jobs that are on hosts we can bill for.
    nonbillable_node_job_details = []  # Jobs not on hosts we can bill for.
    unknown_user_job_details     = []  # Jobs from users we don't know.

    for line in accounting_fp:

        if line[0] == "#": continue

        fields = line.split(':')

        accounting_record = dict(zip(ACCOUNTING_FIELDS, fields))

        # If the job failed, the submission_time is the job date.
        # Else, the end_time is the job date.
        failed_code = int(accounting_record['failed'])
        job_failed = failed_code in ACCOUNTING_FAILED_CODES
        if job_failed:
            job_date = int(accounting_record['submission_time'])  # The only valid date in the record.
        else:
            job_date = int(accounting_record['end_time'])

        # Create a list of job details for this job.
        job_details = []
        job_details.append(job_date)
        job_details.append(job_date)  # Two columns used for the date: one date formatted, one timestamp.
        job_details.append(accounting_record['owner'])
        job_details.append(accounting_record['job_name'])

        # Edit out the default account 'sge'.
        if accounting_record['account'] != 'sge':

            # If this account/job tag is unknown, save details for later output.
            if (accounting_record['account'] not in job_tag_list and
                accounting_record['account'] not in pi_tag_list):
                not_in_job_tag_list.add((accounting_record['owner'],
                                         accounting_record['job_name'],
                                         accounting_record['account']))

            job_details.append(accounting_record['account'])
        else:
            job_details.append('')

        # Edit hostname to remove trailing ".local".
        node_name = accounting_record['hostname']
        if node_name.endswith(".local"):
            node_name = node_name[:-6]

        job_details.append(node_name)

        job_details.append(int(accounting_record['slots']))
        job_details.append(int(accounting_record['ru_wallclock']))
        job_details.append(int(accounting_record['job_number']))

        # If the end date of this job was within the month,
        #  examine it.
        if begin_timestamp <= job_date < end_timestamp:

            # Do we know this job's user?
            if accounting_record['owner'] in users_list:

                # If job failed, save in Failed job list.
                if job_failed:
                    failed_job_details.append(job_details + [failed_code])
                else:
                    # Job is billable if it ran on hosts starting with one of the BILLABLE_HOSTNAME_PREFIXES.
                    billable_hostname_prefixes = map(lambda p: node_name.startswith(p), BILLABLE_HOSTNAME_PREFIXES)

                    # If hostname doesn't have a billable prefix, save in an nonbillable list.
                    if not any(billable_hostname_prefixes):
                        nonbillable_node_job_details.append(job_details + ['Nonbillable Node'])
                    else:
                        billable_job_details.append(job_details)

            else:
                # Save unknown user and job details in unknown user lists.
                not_in_users_list.add(accounting_record['owner'])
                unknown_user_job_details.append(job_details + ['Unknown User'])

        else:
            if job_date != 0:
                dates_tuple = (from_timestamp_to_date_string(job_date),
                               from_timestamp_to_date_string(begin_timestamp),
                               from_timestamp_to_date_string(end_timestamp))
                print "Job date %s is not between %s and %s" % dates_tuple
            else:
                print "Job date is zero."
            print ':'.join(fields)

    # Close the accounting file.
    accounting_fp.close()

    # Print out list of users who had jobs but were not in any lab list.
    if len(not_in_users_list) > 0:
        print "  *** Job submitters not in users list:",
        for user in not_in_users_list:
            print user,
        print
    if len(not_in_job_tag_list) > 0:
        print "  *** Jobs with unknown job tags:"
        for (user, job_name, job_tag) in not_in_job_tag_list:
            print '   ', user, job_name, job_tag

    # Output the accounting details to the BillingDetails worksheet.
    print "  Outputting accounting details"

    # Output jobs to sheet for billable jobs.
    print "    Billable Jobs:    ",
    write_job_details(computing_sheet, billable_job_details)

    # Output nonbillable jobs to sheet for nonbillable jobs.
    print "    Nonbillable Jobs: ",
    all_nonbillable_job_details = nonbillable_node_job_details + unknown_user_job_details
    write_job_details(nonbillable_job_sheet, all_nonbillable_job_details)

    # Output jobs to sheet for failed jobs.
    print "    Failed Jobs:      ",
    write_job_details(failed_job_sheet, failed_job_details)

    print "Computing charges computed."

# Generates the "Consulting" sheet (someday).
def compute_consulting_charges(config_wkbk, begin_timestamp, end_timestamp, consulting_sheet):
    pass

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
parser.add_argument("-r", "--billing_root",
                    default=None,
                    help='The Billing Root directory [default = None]')
parser.add_argument("-v", "--verbose", action="store_true",
                    default=False,
                    help='Get real chatty [default = false]')
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
parser.add_argument("--scg3", action="store_true",
                    default=False,
                    help="Add SCG3 nodes to output [default = false]")
parser.add_argument("-y","--year", type=int, choices=range(2013,2031),
                    default=None,
                    help="The year to be filtered out. [default = this year]")
parser.add_argument("-m", "--month", type=int, choices=range(1,13),
                    default=None,
                    help="The month to be filtered out. [default = last month]")

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

# Add the SCG3 prefix to the filter for hostnames.
if args.scg3:
    BILLABLE_HOSTNAME_PREFIXES.append('scg3')

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

# Use switch arg for accounting_file if present, else use file in BillingRoot.
if args.accounting_file is not None:
    accounting_file = args.accounting_file
else:
    accounting_filename = "%s.%d-%02d.txt" % (SGEACCOUNTING_PREFIX, year, month)
    accounting_file = os.path.join(year_month_dir, accounting_filename)

# Initialize the BillingDetails spreadsheet.
details_wkbk_filename = "%s.%s-%02d.xlsx" % (BILLING_DETAILS_PREFIX, year, month)
details_wkbk_pathname = os.path.join(year_month_dir, details_wkbk_filename)

billing_details_wkbk = xlsxwriter.Workbook(details_wkbk_pathname)
sheet_name_to_sheet = init_billing_details_wkbk(billing_details_wkbk)

#
# Output the state of arguments.
#
print "GETTING DETAILS FOR %02d/%d:" % (month, year)
print "  BillingConfigFile: %s" % (args.billing_config_file)
print "  BillingRoot: %s" % billing_root
print "  SGEAccountingFile: %s" % accounting_file
if args.no_storage:
    print "  Skipping storage calculations"
if args.no_computing:
    print "  Skipping computing calculations"
if args.no_consulting:
    print "  Skipping consulting calculations"
if args.scg3:
    print "  Including scg3 hosts."
print "  BillingDetailsFile: %s" % (details_wkbk_pathname)
print

#
# Compute storage charges.
#
if not args.no_storage:
    compute_storage_charges(billing_config_wkbk, begin_month_timestamp, end_month_timestamp,
                            sheet_name_to_sheet['Storage'])

#
# Compute computing charges.
#
if not args.no_computing:
    compute_computing_charges(billing_config_wkbk, begin_month_timestamp, end_month_timestamp, accounting_file,
                              sheet_name_to_sheet['Computing'],
                              sheet_name_to_sheet['Nonbillable Jobs'], sheet_name_to_sheet['Failed Jobs'])

#
# Compute consulting charges.
#
if not args.no_consulting:
    compute_consulting_charges(billing_config_wkbk, begin_month_timestamp, end_month_timestamp,
                               sheet_name_to_sheet['Consulting'])

#
# Close the output workbook and write the .xlsx file.
#
print "Closing BillingDetails workbook."
billing_details_wkbk.close()