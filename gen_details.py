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
#   --all_jobs_billable: Consider all jobs to be billable. [default=False]
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
import collections
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
USAGE_EXECUTABLE = ['du', '-s']
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
global remove_unicode_chars

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

# Deduces the fileset that a folder being measured by quota lives in.
def get_device_and_fileset_from_folder(folder):

    # Find the two path elements after "/srv/gsfs0/".
    path_elts = os.path.normpath(folder).split(os.path.sep)

    if (len(path_elts) >= 4 and
        path_elts[0] == '' and
        path_elts[1] == 'srv'):

        return (path_elts[2], ".".join(path_elts[3:5]))
    else:
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
        print >> sys.stderr, "ERROR: No fileset for folder %s; ignoring..."
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

# Given a list of tuples of job details, writes the job details to
# the sheet given.  Possible sheets for use in this method are
# typically the "Computing", "Nonbillable Jobs", and "Failed Jobs" sheets.
def write_job_details(sheet, job_details):

    # If no job details, write "No Jobs".
    if len(job_details) == 0:
        sheet.write(1, 0, "No jobs")
        print
        return

    # If we have job details, write them to the sheet, below the headers.
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
        sheet.write(sheet_row, col, unicode(job_details[row][col],'utf-8'))
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

        # 'Wallclock Secs'
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
    quota_bools = sheet_get_named_column(folder_sheet, 'Method')
    dates_added = sheet_get_named_column(folder_sheet, 'Date Added')
    dates_remvd = sheet_get_named_column(folder_sheet, 'Date Removed')

    # Mapping from folders to [timestamp, total, used].
    folder_size_dict = collections.OrderedDict()

    # Create mapping from folders to space used.
    for (folder, pi_tag, quota_bool, date_added, date_removed) in zip(folders, pi_tags, quota_bools, dates_added, dates_remvd):

        # Skip measuring this folder if we have already done it.
        if folder_size_dict.get(folder) is not None: continue

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

            folder_size_dict[folder] = [time.time(), total, used]
        else:
            print "  *** Excluding %s for PI %s: folder not active in this month" % (folder, pi_tag)

    # Write space-used mapping into details workbook.
    row = 0
    for folder in folder_size_dict.keys():

        [timestamp, total, used] = folder_size_dict[folder]
        sheet_row = row + 1

        # 'Date Measured'
        col = 0
        storage_sheet.write(sheet_row, col, from_timestamp_to_excel_date(timestamp), DATE_FORMAT)

        # 'Timestamp'
        col += 1
        storage_sheet.write(sheet_row, col, timestamp)
        if args.verbose: print timestamp,

        # 'Folder'
        col += 1
        storage_sheet.write(sheet_row, col, folder)
        if args.verbose: print folder,

        # 'Size'
        col += 1
        storage_sheet.write(sheet_row, col, total, FLOAT_FORMAT)
        if args.verbose: print total,

        # 'Used'
        col += 1
        storage_sheet.write(sheet_row, col, used, FLOAT_FORMAT)
        if args.verbose: print used

        # Next row, please.
        row += 1


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
    not_in_job_tag_list = collections.defaultdict(set)
    both_proj_and_acct_list = collections.defaultdict(set)

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

        #
        # Look for job tags in both account and project fields.
        # If values occur in both, use the project field and record the discrepancy.
        #
        account = remove_unicode_chars(accounting_record['account'])
        if account == 'sge':   # Edit out the default account 'sge'.
            account = None

        project = remove_unicode_chars(accounting_record['project'])
        if project == 'NONE':  # Edit out the placeholder project 'NONE'.
            project = None

        #
        # Add job tag (project/account) info to job_details.
        #
        job_tag = None
        if project is not None:

            # The project is a valid job tag if it is either in the job_tag_list
            #  or the pi_tag_list.
            project_is_valid_job_tag = (project in job_tag_list or project in pi_tag_list)

            if not project_is_valid_job_tag:
                # If this project/job tag is unknown, save details for later output.
                not_in_job_tag_list[accounting_record['owner']].add(project)
        else:
            project_is_valid_job_tag = False

        if account is not None:

            # The account is a valid job tag if it is either in the job_tag_list
            #  or the pi_tag_list.
            account_is_valid_job_tag = (account in job_tag_list or account.lower() in pi_tag_list)

            if not account_is_valid_job_tag:
                # If this account/job tag is unknown, save details for later output.
                not_in_job_tag_list[accounting_record['owner']].add(account)

        else:
            account_is_valid_job_tag = False

        # Decide which of project and account will be used for job tag.
        if project_is_valid_job_tag:

            # If there's both a project and an account, choose the project and save details for later output.
            job_tag = project
            if account_is_valid_job_tag:
                both_proj_and_acct_list[accounting_record['owner']].add((project,account))

        elif account_is_valid_job_tag:
            job_tag = account

        # Add the computed job_tag to the job_details, if any.
        if job_tag is not None:
            job_details.append(job_tag)
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

        # If the end date of this job was within the month or we aren't reading job timestamps,
        #  examine it.
        if (args.ignore_job_timestamps or begin_timestamp <= job_date < end_timestamp):

            # Is the job's node billable?
            if not args.all_jobs_billable:
                # Job is billable if it ran on hosts starting with one of the BILLABLE_HOSTNAME_PREFIXES.
                billable_hostname_prefixes = map(lambda p: node_name.startswith(p), BILLABLE_HOSTNAME_PREFIXES)
            else:
                billable_hostname_prefixes = [True]
            job_node_is_billable = any(billable_hostname_prefixes)

            # Do we know this job's user?
            job_user_is_known = accounting_record['owner'] in users_list
            # If not, save the username in an unknown-user list.
            if not job_user_is_known:
                # Save unknown user and job details in unknown user lists.
                not_in_users_list.add(accounting_record['owner'])

            # If we know the user or the job has a job tag,...
            if job_user_is_known or job_tag is not None:

                # If job failed, save in Failed job list.
                if job_failed:
                    failed_job_details.append(job_details + [failed_code])
                else:
                    # If hostname doesn't have a billable prefix, save in an nonbillable list.
                    if not job_node_is_billable:
                        nonbillable_node_job_details.append(job_details + ['Nonbillable Node'])
                    else:
                        billable_job_details.append(job_details)

            else:
                # Save the job details in an unknown-user list.
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
    # Print out list of unknown job tags.
    if len(not_in_job_tag_list.keys()) > 0:
        print "  *** Jobs with unknown job tags:"
        for user in sorted(not_in_job_tag_list.keys()):
            print '   ', user
            for job_tag in sorted(not_in_job_tag_list[user]):
                print '     ', job_tag
    # Print out list of jobs with both project and account job tags.
    if len(both_proj_and_acct_list.keys()) > 0:
        print "  *** Jobs with both project and account job tags:"
        for user in sorted(both_proj_and_acct_list.keys()):
            print '   ', user
            for (proj, acct) in both_proj_and_acct_list[user]:
                print '     Project:', proj, 'Account:', acct

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
parser.add_argument("--all_jobs_billable", action="store_true",
                    default=False,
                    help="Consider all jobs to be billable [default = false]")
parser.add_argument("-i", "--ignore_job_timestamps", action="store_true",
                    default=False,
                    help="Ignore timestamps in job (and allow jobs not in month selected) [default = false]")
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
if args.all_jobs_billable:
    print "  All jobs billable."
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