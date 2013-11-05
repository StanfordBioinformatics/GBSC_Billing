#!/usr/bin/env python

#===============================================================================
#
# gen_details.py - Generate billing details for month/year.
#
# ARGS:
#   1st: the BillingConfig spreadsheet.
#
# SWITCHES:
#
# OUTPUT:
#
# ASSUMPTIONS:
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

#=====
#
# CONSTANTS
#
#=====
BILLING_DETAILS_PREFIX = "BillingDetails"

# Mapping from sheet name to the column headers within that sheet.
BILLING_DETAILS_SHEET_COLUMNS = {
    'Storage'    : ('Date Measured', 'Timestamp', 'Folder', 'Size', 'Used'),
    'Computing'  : ('Job Date', 'Job Timestamp', 'Username', 'Job Name', 'Account', 'Node', 'Cores', 'Wallclock', 'Job ID'),
    'Consulting' : ('Work Date', 'Item', 'Hours', 'PI')
}

QUOTA_EXECUTABLE = ['/usr/lpp/mmfs/bin/mmlsquota', '-j']
USAGE_EXECUTABLE = ['sudo', 'du', '-s']
STORAGE_BLOCK_SIZE_ARG = ['--block-size=1G']  # Works in both above commands.

# OGE accounting file column info: 
# http://manpages.ubuntu.com/manpages/lucid/man5/sge_accounting.5.html
ACCOUNTING_FIELDS = (
    'qname', 'hostname', 'group', 'owner', 'job_name', 'job_number',
    'account', 'priority','submission_time', 'start_time', 'end_time',
    'failed', 'exit_status', 'ru_wallclock', 'ru_utime', 'ru_stime',
    'ru_maxrss', 'ru_ixrss', 'ru_ismrss', 'ru_idrss', 'ru_isrss', 'ru_minflt', 'ru_majflt',
    'ru_nswap', 'ru_inblock', 'ru_oublock', 'ru_msgsnd', 'ru_msgrcv', 'ru_nsignals',
    'ru_nvcsw', 'ru_nivcsw', 'project', 'department', 'granted_pe', 'slots',
    'task_number', 'cpu', 'mem', 'category', 'iow', 'pe_taskid', 'max_vmem', 'arid',
    'ar_submission_time'
)

SGEACCOUNTING_PREFIX = "SGEAccounting"  # Prefix of accounting file name in BillingRoot.

# List of hostname prefixes to use for billing purposes.
HOSTNAME_FILTER_PREFIXES = ['scg1']

#=====
#
# GLOBALS
#
#=====

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

# This method takes in an xlrd Sheet object and a column name,
# and returns all the values from that column.
def sheet_get_named_column(sheet, col_name):

    header_row = sheet.row_values(0)

    for idx in range(len(header_row)):
        if header_row[idx] == col_name:
           col_name_idx = idx
           break
    else:
        return None

    return sheet.col_values(col_name_idx,start_rowx=1)


def config_sheet_get_dict(wkbk):

    config_sheet = wkbk.sheet_by_name("Config")

    config_keys   = sheet_get_named_column(config_sheet, "Key")
    config_values = sheet_get_named_column(config_sheet, "Value")

    return dict(zip(config_keys, config_values))


def read_config_sheet(wkbk):

    config_dict = config_sheet_get_dict(wkbk)

    accounting_file = config_dict.get("SGEAccountingFile")
    if accounting_file is None:
        print >> sys.stderr, "Need accounting file: exiting..."
        sys.exit(-1)

    billing_root    = config_dict.get("BillingRoot", os.getcwd())

    return (billing_root, accounting_file)


def init_billing_details_wkbk(workbook):

    global BOLD_FORMAT
    global DATE_FORMAT
    global FLOAT_FORMAT
    global INT_FORMAT
    global MONEY_FORMAT
    global PERCENT_FORMAT

    # Create formats for use within the workbook.
    BOLD_FORMAT = workbook.add_format({'bold' : True})
    DATE_FORMAT = workbook.add_format({'num_format' : 'mm/dd/yy'})
    INT_FORMAT  = workbook.add_format({'num_format' : '0'})
    FLOAT_FORMAT = workbook.add_format({'num_format' : '0.0'})
    MONEY_FORMAT = workbook.add_format({'num_format' : '$0.00'})
    PERCENT_FORMAT = workbook.add_format({'num_format' : '0%'})

    sheet_name_to_sheet = dict()

    for sheet_name in BILLING_DETAILS_SHEET_COLUMNS.keys():

        sheet = workbook.add_worksheet(sheet_name)
        for col in range(0, len(BILLING_DETAILS_SHEET_COLUMNS[sheet_name])):
            sheet.write(0, col, BILLING_DETAILS_SHEET_COLUMNS[sheet_name][col], BOLD_FORMAT)

        sheet_name_to_sheet[sheet_name] = sheet

    # Make the Storage sheet the active one.
    sheet_name_to_sheet['Storage'].activate()

    return sheet_name_to_sheet


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

        #print begin_timestamp, date_added, end_timestamp
        #if begin_timestamp <= date_added:
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

    # Write space-used mapping into details workbook.
    for row in range(0, len(folder_sizes)):

        sheet_row = row + 1

        # 'Date Measured'
        col = 0
        storage_sheet.write_formula(sheet_row, col, '=(%f/86400)+DATE(1970,1,1)' % folder_sizes[row][col], DATE_FORMAT)

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


def compute_computing_charges(config_wkbk, begin_timestamp, end_timestamp, accounting_file, computing_sheet):

    print "Computing computing charges..."

    #
    # Read in the Usernames from the Users sheet.
    #
    users_sheet = config_wkbk.sheet_by_name('Users')

    users_list = sheet_get_named_column(users_sheet, "Username")
    #  NOTE: This column may have some duplicates in it.
    #        Need to make a set out of the result.
    users = set(users_list)

    #
    # Read in the Job Tags from the Job Tags sheet.
    #
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
    #  which have "submission_times" in the given month,
    #  and "owner"s in the list of users.
    #
    not_in_users_list = set()
    not_in_job_tag_list = set()

    accounting_details = []
    for line in accounting_fp:

        if line[0] == "#": continue

        fields = line.split(':')

        accounting_record = dict(zip(ACCOUNTING_FIELDS, fields))

        #
        # Only look at hosts that start with "scg1".
        #

        # Edit hostname to remove trailing ".local".
        node_name = accounting_record['hostname']
        if node_name.endswith(".local"):
            node_name = node_name[:-6]

        filtering_prefixes = map(lambda p: node_name.startswith(p), HOSTNAME_FILTER_PREFIXES)
        if not any(filtering_prefixes): continue

        submission_date = int(accounting_record['submission_time'])

        # If the submission date of this job was within the month,
        #  examine it.
        if begin_timestamp <= submission_date < end_timestamp:

            if accounting_record['owner'] in users:

                job_details = []
                job_details.append(submission_date)
                job_details.append(submission_date)  # Two columns used for the date: one date formatted, one timestamp.
                job_details.append(accounting_record['owner'])
                job_details.append(accounting_record['job_name'])

                # Elide the default account 'sge'.
                if accounting_record['account'] != 'sge':

                    # If this account/job tag is unknown, save details for later output.
                    if accounting_record['account'] not in job_tag_list:
                        not_in_job_tag_list.add((accounting_record['owner'],
                                                 accounting_record['job_name'],
                                                 accounting_record['account']))

                    job_details.append(accounting_record['account'])
                else:
                    job_details.append('')

                job_details.append(node_name)

                job_details.append(int(accounting_record['slots']))
                job_details.append(int(accounting_record['ru_wallclock']))
                job_details.append(int(accounting_record['job_number']))

                accounting_details.append(job_details)

            else:
                not_in_users_list.add(accounting_record['owner'])

        else:
            dates_tuple = (datetime.date.fromtimestamp(submission_date).strftime("%m/%d/%Y"),
                           datetime.date.fromtimestamp(begin_timestamp).strftime("%m/%d/%Y"),
                           datetime.date.fromtimestamp(end_timestamp).strftime("%m/%d/%Y"))
            print "Job date %s is not between %s and %s" % dates_tuple

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

    for row in range(0, len(accounting_details)):

        # Bump rows down below header line.
        sheet_row = row + 1

        # A little feedback for the people.
        if not args.verbose:
            if sheet_row % 1000 == 0:
                sys.stdout.write('.')
                sys.stdout.flush()

        # 'Job Date'
        col = 0
        computing_sheet.write_formula(sheet_row, col, '=(%f/86400)+DATE(1970,1,1)' % accounting_details[row][col], DATE_FORMAT)

        # 'Job Timestamp'
        col += 1
        computing_sheet.write(sheet_row, col, accounting_details[row][col])
        if args.verbose: print accounting_details[row][col],

        # 'Username'
        col += 1
        computing_sheet.write(sheet_row, col, accounting_details[row][col])
        if args.verbose: print accounting_details[row][col],

        # 'Job Name'
        col += 1
        computing_sheet.write(sheet_row, col, accounting_details[row][col])
        if args.verbose: print accounting_details[row][col],

        # 'Account'
        col += 1
        computing_sheet.write(sheet_row, col, accounting_details[row][col])
        if args.verbose: print accounting_details[row][col],

        # 'Node'
        col += 1
        computing_sheet.write(sheet_row, col, accounting_details[row][col])
        if args.verbose: print accounting_details[row][col],

        # 'Slots'
        col += 1
        computing_sheet.write(sheet_row, col, accounting_details[row][col], INT_FORMAT)
        if args.verbose: print accounting_details[row][col],

        # 'Wallclock'
        col += 1
        computing_sheet.write(sheet_row, col, accounting_details[row][col], INT_FORMAT)
        if args.verbose: print accounting_details[row][col],

        # 'JobID'
        col += 1
        computing_sheet.write(sheet_row, col, accounting_details[row][col], INT_FORMAT)
        if args.verbose: print accounting_details[row][col],

        if args.verbose: print

    print

    print "Computing charges computed."


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
                    help='The SGE accounting file to snapshot [default = None]')
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
# Both values should be GMT.
begin_month_timestamp = int(time.mktime(datetime.date(year, month, 1).timetuple()))
end_month_timestamp   = int(time.mktime(datetime.date(next_month_year, next_month, 1).timetuple()))

# Add the SCG3 prefix to the filter for hostnames.
if args.scg3:
    HOSTNAME_FILTER_PREFIXES.append('scg3')

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
    compute_storage_charges(billing_config_wkbk, begin_month_timestamp, end_month_timestamp, sheet_name_to_sheet['Storage'])

#
# Compute computing charges.
#
if not args.no_computing:
    compute_computing_charges(billing_config_wkbk, begin_month_timestamp, end_month_timestamp, accounting_file, sheet_name_to_sheet['Computing'])

#
# Compute consulting charges.
#
if not args.no_consulting:
    compute_consulting_charges(billing_config_wkbk, begin_month_timestamp, end_month_timestamp, sheet_name_to_sheet['Consulting'])

#
# Close the output workbook and write the .xlsx file.
#
print "Closing BillingDetails workbook."
billing_details_wkbk.close()