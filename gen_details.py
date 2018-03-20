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
#   --ignore_job_timestamps: Ignore timestamps in job and allow jobs not in month selected [default=False]
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
import codecs
import collections
import csv
import datetime
import locale
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
global GOOGLE_INVOICE_PREFIX
global BILLING_DETAILS_PREFIX
global CONSULTING_PREFIX
global STORAGE_PREFIX

#
# Formats for the output workbook, to be initialized along with the workbook.
#
BOLD_FORMAT = None
DATE_FORMAT = None
INT_FORMAT  = None
FLOAT_FORMAT = None
MONEY_FORMAT = None
PERCENT_FORMAT = None

# Set locale to be US english for converting strings with commas into floats.
locale.setlocale(locale.LC_ALL, 'en_US.UTF-8')

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
global filter_by_dates

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

    sheet_name_to_sheet_map = dict()

    for sheet_name in BILLING_DETAILS_SHEET_COLUMNS:

        sheet = workbook.add_worksheet(sheet_name)
        for col in range(0, len(BILLING_DETAILS_SHEET_COLUMNS[sheet_name])):
            sheet.write(0, col, BILLING_DETAILS_SHEET_COLUMNS[sheet_name][col], BOLD_FORMAT)

        sheet_name_to_sheet_map[sheet_name] = sheet

    # Make the Storage sheet the active one.
    sheet_name_to_sheet_map['Storage'].activate()

    return sheet_name_to_sheet_map


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

#
# Reads a subtable from the CSVFile fileobject, which is all the lines
# between blank lines.
#
def get_google_invoice_csv_subtable_lines(csvfile_obj):

    subtable = []

    line = csvfile_obj.readline()
    while line != '' and line != '\n':
        subtable.append(line)
        line = csvfile_obj.readline()

    return subtable


# Read the Storage Usage file.
# Returns mapping from folders to [timestamp, total, used]
def read_storage_usage_file(storage_usage_file):

    # Mapping from folders to [timestamp, total, used].
    folder_size_dict = collections.OrderedDict()

    usage_fileobj = open(storage_usage_file)

    usage_csvdict = csv.DictReader(usage_fileobj)

    for row in usage_csvdict:
        folder_size_dict[row['Folder']] = (float(row['Timestamp']), float(row['Size']), float(row['Used']))

    return folder_size_dict


# Write storage usage data into the Storage sheet of the BillingDetails file.
# Takes in mapping from folders to [timestamp, total, used].
def write_storage_usage_data(folder_size_dict, storage_sheet):

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
    global NONBILLABLE_HOSTNAME_PREFIXES
    global IGNORED_JOB_TAGS

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
    unknown_node_job_details     = []  # Jobs on unknown nodes.
    unknown_user_job_details     = []  # Jobs from users we don't know.

    unknown_job_nodes            = set()  # Set of nodes we don't know.

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
        if account == 'sge' or account == '':   # Edit out the default account 'sge'.
            account = None

        project = remove_unicode_chars(accounting_record['project'])
        if project == 'NONE' or project == '':  # Edit out the placeholder project 'NONE'.
            project = None

        #
        # Add job tag (project/account) info to job_details.
        #
        # If project is set and not in the ignored job tag list:
        if project is not None and project not in IGNORED_JOB_TAGS:

            # The project is a valid job tag if it is either in the job_tag_list
            #  or the pi_tag_list.
            project_is_valid_job_tag = (project in job_tag_list or project in pi_tag_list)

            if not project_is_valid_job_tag:
                # If this project/job tag is unknown, save details for later output.
                not_in_job_tag_list[accounting_record['owner']].add(project)
        else:
            project_is_valid_job_tag = False
            project = None  # we could be ignoring a given job tag

        # If account is set and not in the ignored job tag list:
        if account is not None and account not in IGNORED_JOB_TAGS:

            # The account is a valid job tag if it is either in the job_tag_list
            #  or the pi_tag_list.
            account_is_valid_job_tag = \
                (account in job_tag_list or
                 account.lower() in job_tag_list or
                 account.lower() in pi_tag_list)

            if not account_is_valid_job_tag:
                # If this account/job tag is unknown, save details for later output.
                not_in_job_tag_list[accounting_record['owner']].add(account)
        else:
            account_is_valid_job_tag = False
            account = None  # we could be ignoring a given job tag

        #
        # Decide which of project and account will be used for job tag.
        #

        # If project is valid, choose project for job tag.
        if project_is_valid_job_tag:

            # If there's both a project and an account, choose the project and save details for later output.
            job_tag = project
            if account is not None:
                both_proj_and_acct_list[accounting_record['owner']].add((project,account))

        # Else if project is present and account is not valid, choose project for job tag.
        # (Non-valid project trumps non-valid account).
        elif project is not None and not account_is_valid_job_tag:

            # If there's both a project and an account, choose the project and save details for later output.
            job_tag = project
            if account is not None:
                both_proj_and_acct_list[accounting_record['owner']].add((project,account))

        # Else if account is present, choose account for job tag.
        # (either account is valid and the project is non-valid, or there is no project).
        elif account is not None:
            job_tag = account

            # If there's both an account and a project, save the details for later output.
            if project is not None:
                both_proj_and_acct_list[accounting_record['owner']].add((project,account))

        # else No project and No account = No job tag.
        else:
            job_tag = None

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
                nonbillable_hostname_prefixes = map(lambda p: node_name.startswith(p), NONBILLABLE_HOSTNAME_PREFIXES)
            else:
                billable_hostname_prefixes = [True]
                nonbillable_hostname_prefixes = []

            job_node_is_billable = any(billable_hostname_prefixes)
            job_node_is_nonbillable = any(nonbillable_hostname_prefixes)

            job_node_is_unknown_billable = not (job_node_is_billable or job_node_is_nonbillable)

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
                    if job_node_is_billable:
                        billable_job_details.append(job_details)
                    elif job_node_is_nonbillable:
                        nonbillable_node_job_details.append(job_details + ['Nonbillable Node'])
                    elif job_node_is_unknown_billable:
                        unknown_node_job_details.append(job_details + ['Unknown Node'])
                        unknown_job_nodes.add(node_name)
                    else:
                        pass # SHOULD NOT GET HERE.
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

    #
    # ERROR FLAGGING:
    #

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
    # Print out how many jobs were run on unknown nodes.
    if len(unknown_job_nodes) > 0:
        print "  *** Unknown Nodes with jobs:"
        for node in sorted(unknown_job_nodes):
            print '   ', node

    # Output the accounting details to the BillingDetails worksheet.
    print "  Outputting accounting details"

    # Output jobs to sheet for billable jobs.
    if len(billable_job_details) > 0:
        print "    Billable Jobs:    ",
        write_job_details(computing_sheet, billable_job_details)

    # Output nonbillable jobs to sheet for nonbillable jobs.
    if len(nonbillable_node_job_details) > 0:
        print "    Nonbillable Jobs: ",
        all_nonbillable_job_details = \
            nonbillable_node_job_details + unknown_user_job_details + unknown_node_job_details
        write_job_details(nonbillable_job_sheet, all_nonbillable_job_details)

    # Output jobs to sheet for failed jobs.
    if len(failed_job_details) > 0:
        print "    Failed Jobs:      ",
        write_job_details(failed_job_sheet, failed_job_details)

    print "Computing charges computed."

# Generates the "Consulting" sheet.
def compute_consulting_charges(config_wkbk, begin_timestamp, end_timestamp, consulting_timesheet, consulting_sheet):

    print "Computing consulting charges..."

    ###
    # Read the config workbook to get a list of active PIs
    ###
    pis_sheet = config_wkbk.sheet_by_name("PIs")

    pis_list    = sheet_get_named_column(pis_sheet, "PI Tag")
    dates_added = sheet_get_named_column(pis_sheet, "Date Added")
    dates_remvd = sheet_get_named_column(pis_sheet, "Date Removed")

    active_pis_list = filter_by_dates(pis_list, zip(dates_added, dates_remvd), begin_timestamp, end_timestamp)

    ###
    # Read the Consulting Timesheet.
    ###

    consulting_workbook = xlrd.open_workbook(consulting_timesheet)

    hours_sheet = consulting_workbook.sheet_by_name("Hours")

    dates   = sheet_get_named_column(hours_sheet, "Date")
    pi_tags = sheet_get_named_column(hours_sheet, "PI Tag")
    hours   = sheet_get_named_column(hours_sheet, "Hours")
    travel_hours = sheet_get_named_column(hours_sheet, "Travel Hours")
    participants = sheet_get_named_column(hours_sheet, "Participants")
    summaries = sheet_get_named_column(hours_sheet, "Summary")
    notes   = sheet_get_named_column(hours_sheet, "Notes")
    cumul_hours = sheet_get_named_column(hours_sheet, "Cumul Hours")

    # Convert empty travel hours to zeros.
    travel_hours = map(lambda h: 0 if h=='' else h, travel_hours)

    row = 1
    for (date, pi_tag, hours_spent, travel_hrs, participant, summary, note, cumul_hours_spent) in \
            zip(dates, pi_tags, hours, travel_hours, participants, summaries, notes, cumul_hours):

        # Confirm date is within this month.
        date_timestamp = from_excel_date_to_timestamp(date)

        # If date is before beginning of the month or after the end of the month, skip this entry.
        if begin_timestamp > date_timestamp or date_timestamp >= end_timestamp: continue

        # Confirm that pi_tag is in the list of active PIs for this month.
        if pi_tag not in active_pis_list:
            print "  PI %s not in active PI list...skipping" % pi_tag

        # Copy the entry into the output consulting sheet.
        col = 0
        consulting_sheet.write(row, col, date, DATE_FORMAT)
        col += 1
        consulting_sheet.write(row, col, pi_tag)
        col += 1
        consulting_sheet.write(row, col, float(hours_spent), FLOAT_FORMAT)
        col += 1
        consulting_sheet.write(row, col, float(travel_hrs), FLOAT_FORMAT)
        col += 1
        consulting_sheet.write(row, col, participant)
        col += 1
        consulting_sheet.write(row, col, summary)
        col += 1
        consulting_sheet.write(row, col, note)
        col += 1
        consulting_sheet.write(row, col, float(cumul_hours_spent), FLOAT_FORMAT)
        col += 1

        row += 1


def write_cloud_details_V1(cloud_sheet, row_dict, output_row):

    output_col = 0
    total_amount = 0.0

    # Write Google data into Cloud sheet.

    # Output 'Platform' field.
    cloud_sheet.write(output_row, output_col, row_dict['Product'])
    output_col += 1

    # Output 'Account' Field.
    cloud_sheet.write(output_row, output_col, row_dict['Order'])
    output_col += 1

    # Output 'Project' field.
    cloud_sheet.write(output_row, output_col, row_dict['Source'])
    output_col += 1

    # Output 'Description' field.
    cloud_sheet.write(output_row, output_col, row_dict['Description'])
    output_col += 1

    # Output 'Dates' field.
    cloud_sheet.write(output_row, output_col, row_dict['Interval'])
    output_col += 1

    # Parse quantity.
    if len(row_dict['Quantity']) > 0:
        quantity = locale.atof(row_dict['Quantity'])
    else:
        quantity = ''

    # Output 'Quantity' field.
    cloud_sheet.write(output_row, output_col, quantity, FLOAT_FORMAT)
    output_col += 1

    # Output 'Unit of Measure' field.
    cloud_sheet.write(output_row, output_col, row_dict['UOM'])
    output_col += 1

    # Parse charge.
    amount = locale.atof(row_dict['Amount'])
    # Accumulate total charges.
    total_amount += amount

    # Output 'Charge' field.
    cloud_sheet.write(output_row, output_col, amount, MONEY_FORMAT)
    output_col += 1

    return total_amount


def write_cloud_details_V2(cloud_sheet, row_dict, output_row):

    output_col = 0
    total_amount = 0.0

    # Write Google data into Cloud sheet.

    # Output 'Platform' field.
    cloud_sheet.write(output_row, output_col, "Google Cloud Platform")
    output_col += 1

    # Output 'Account' field. (subacccount)
    cloud_sheet.write(output_row, output_col, row_dict['Account ID'])
    output_col += 1

    # Output 'Project' field.  (Project Name + Project ID)
    cloud_sheet.write(output_row, output_col, row_dict['Source'])
    output_col += 1

    # Output 'Description' field. (SKU description of the charge)
    sku_description = "%s %s" % (row_dict['Product'], row_dict['Resource Type'])
    cloud_sheet.write(output_row, output_col, sku_description)
    output_col += 1

    # Output 'Dates' field.
    date_range = "%s-%s" % (row_dict['Start Date'], row_dict['End Date'])
    cloud_sheet.write(output_row, output_col, date_range)
    output_col += 1

    # Parse quantity.
    quantity_str = row_dict['Quantity'].strip()
    if len(quantity_str) > 0:
        quantity = locale.atof(quantity_str)
    else:
        quantity = ''

    # Output 'Quantity' field.
    cloud_sheet.write(output_row, output_col, quantity, FLOAT_FORMAT)
    output_col += 1

    # Output 'Unit of Measure' field.
    cloud_sheet.write(output_row, output_col, row_dict['Unit'])
    output_col += 1

    # Parse charge.
    amount = locale.atof(row_dict['Amount'])
    # Accumulate total charges.
    total_amount += amount

    # Output 'Charge' field.
    cloud_sheet.write(output_row, output_col, amount, MONEY_FORMAT)
    output_col += 1

    return total_amount


# Generates the "Cloud" sheet.
def compute_cloud_charges(config_wkbk, google_invoice_csv, cloud_sheet):

    print "Computing cloud charges..."

    ###
    # Read the Google Invoice CSV File
    ###

    # Google Invoice CSV files are Unicode with BOM.
    google_invoice_csv_file_obj = codecs.open(google_invoice_csv, 'rU', encoding='utf-8-sig')

    #  Read the header subtable
    google_invoice_header_subtable = get_google_invoice_csv_subtable_lines(google_invoice_csv_file_obj)

    google_invoice_header_csvreader = csv.DictReader(google_invoice_header_subtable, fieldnames=['key','value'])

    # Determine version of Google CSV file from header subtable.
    # Version 1: has keys "Issue date" and "Amount due".
    # Version 2: has keys "Invoice date" and no "Amount due".
    google_invoice_version = None
    for row in google_invoice_header_csvreader:

        #
        #   Extract invoice date from "Issue Date" or "Invoice date".
        #
        if row['key'] == 'Issue date':
            google_invoice_issue_date = row['value']
            google_invoice_version = 'V1'

        elif row['key'] == 'Invoice date':
            google_invoice_issue_date = row['value']
            google_invoice_version = 'V2'

        #
        #   Extract the "Amount Due" value.
        #
        elif row['key'] == 'Amount due':
            google_invoice_amount_due = locale.atof(row['value'])
            google_invoice_version = "V1"

        elif row['key'] == 'Invoice amount':
            google_invoice_amount_due = locale.atof(row['value'])
            google_invoice_version = "V2"

    if google_invoice_version is None:
        print >> sys.stderr, "  Google Invoice Version not recognized...skipping cloud"
        return

    if args.verbose:
        print >> sys.stderr, "  Google Invoice Version %s" % (google_invoice_version)
        print >> sys.stderr, "  Amount due: $%0.2f" % (google_invoice_amount_due)

    # Accumulate the total amount of charges while processing each line,
    #  to compare with total amount in header in google_invoice_amount_due above.
    google_invoice_total_amount = 0.0

    output_row = 1  # Keeps track of output row in Cloud sheet; starts at 1, below header.

    #  While there are still more subtables...
    while True:

        #   Read subtable.
        google_invoice_subtable = get_google_invoice_csv_subtable_lines(google_invoice_csv_file_obj)

        #   No more subtables?!  Let's get out of here!
        if len(google_invoice_subtable) == 0:
            break

        #   Create CSVReader from subtable
        google_invoice_subtable_csvreader = csv.DictReader(google_invoice_subtable)

        #   Foreach row in CSVReader
        for row_dict in google_invoice_subtable_csvreader:

            if google_invoice_version == 'V1':
                row_amount = write_cloud_details_V1(cloud_sheet, row_dict, output_row)
                if args.verbose: print ".",
            elif google_invoice_version == 'V2':
                row_amount = write_cloud_details_V2(cloud_sheet, row_dict, output_row)
                if args.verbose: print ".",

            # Add up the row charges to compare to total invoice amount.
            google_invoice_total_amount += row_amount

            # Move to next row.
            output_row += 1

    if args.verbose: print

    # Compare total charges to "Amount Due".
    if abs(google_invoice_total_amount - google_invoice_amount_due) >= 0.01:  # Ignore differences less than a penny.
        print >> sys.stderr, "  WARNING: Google accumulated amounts do not equal amount due: ($%.2f != $%.2f)" % (google_invoice_total_amount,
                                                                                                           google_invoice_amount_due)


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
parser.add_argument("-g", "--google_invoice_csv",
                    default=None,
                    help="The Google Invoice CSV file")
parser.add_argument("-c", "--consulting_timesheet",
                    default=None,
                    help="The consulting timesheet XSLX file.")
parser.add_argument("-s", "--storage_usage_csv",
                    default=None,
                    help="The storage usage CSV file.")
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
parser.add_argument("--no_cloud", action="store_true",
                    default=False,
                    help="Don't run cloud calculations [default = false]")
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

# Get absolute path for billing_config_file.
billing_config_file = os.path.abspath(args.billing_config_file)

#
# Open the Billing Config workbook.
#
billing_config_wkbk = xlrd.open_workbook(billing_config_file)

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

# Get absolute path for billing_root directory.
billing_root = os.path.abspath(billing_root)

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

# Get absolute path for accounting_file.
accounting_file = os.path.abspath(accounting_file)

# Use switch arg for google_invoice_file if present, else use file in BillingRoot.
if args.google_invoice_csv is not None:
    google_invoice_csv = args.google_invoice_csv
else:
    google_invoice_filename = "%s.%d-%02d.csv" % (GOOGLE_INVOICE_PREFIX, year, month)
    google_invoice_csv = os.path.join(year_month_dir, google_invoice_filename)

# Get absolute path for google_invoice_csv file.
google_invoice_csv = os.path.abspath(google_invoice_csv)

# Confirm that Google Invoice CSV file exists.
if not os.path.exists(google_invoice_csv):
    google_invoice_csv = None

# Use switch arg to read in storage usage file if given.
#  If not given, generate storage data by analyzing folders.
if args.storage_usage_csv is not None:
    storage_usage_file = args.storage_usage_csv
else:
    storage_usage_filename = "%s.%d-%02d.csv" % (STORAGE_PREFIX, year, month)
    storage_usage_file = os.path.join(year_month_dir, storage_usage_filename)

# Get absolute path for storage_usage_file.
storage_usage_file = os.path.abspath(storage_usage_file)

# Confirm that the storage usage file exists.
if not os.path.exists(storage_usage_file):
    storage_usage_file = None

# Use switch arg for consulting_timesheet if present, else use file in BillingRoot.
if args.consulting_timesheet is not None:
    consulting_timesheet = args.consulting_timesheet
else:
    consulting_filename = "%s.%d-%02d.xlsx" % (CONSULTING_PREFIX, year, month)
    consulting_timesheet = os.path.join(year_month_dir, consulting_filename)

# Get absolute path for consulting_timesheet file.
consulting_timesheet = os.path.abspath(consulting_timesheet)

# Confirm that the consulting_timesheet file exists.
if not os.path.exists(consulting_timesheet):
    consulting_timesheet = None

# Initialize the BillingDetails output spreadsheet.
details_wkbk_filename = "%s.%s-%02d.xlsx" % (BILLING_DETAILS_PREFIX, year, month)
details_wkbk_pathname = os.path.join(year_month_dir, details_wkbk_filename)

billing_details_wkbk = xlsxwriter.Workbook(details_wkbk_pathname)
# Create all the sheets in the output spreadsheet.
sheet_name_to_sheet_map = init_billing_details_wkbk(billing_details_wkbk)

#
# Output the state of arguments.
#
print "GETTING DETAILS FOR %02d/%d:" % (month, year)
print "  BillingConfigFile: %s" % billing_config_file
print "  BillingRoot: %s" % billing_root

if args.no_storage:
    print "  Skipping storage calculations"
    skip_storage = True
elif storage_usage_file is not None:
    print "  Storage usage file: %s" % storage_usage_file
    skip_storage = False
else:
    print "  No storage usage file given...skipping storage calculations"
    skip_storage = True

if args.no_computing:
    print "  Skipping computing calculations"
    skip_computing = True
elif accounting_file is not None:
    print "  SGEAccountingFile: %s" % accounting_file
    skip_computing = False
else:
    print "  No accounting file given...skipping computing calculations"
    skip_computing = True
if args.all_jobs_billable:
    print "  All jobs billable."

if args.no_consulting:
    print "  Skipping consulting calculations"
    skip_consulting = True
elif consulting_timesheet is not None:
    print "  Consulting Timesheet: %s" % consulting_timesheet
    skip_consulting = False
else:
    print "  No consulting timesheet given...skipping consulting calculations"
    skip_consulting = True

if args.no_cloud:
    print "  Skipping cloud calculations"
    skip_cloud = True
elif google_invoice_csv is not None:
    print "  Google Invoice File: %s" % google_invoice_csv
    skip_cloud = False
else:
    print "  No Google Invoice file given...skipping cloud calculations"
    skip_cloud = True

print
print "  BillingDetailsFile to be output: %s" % details_wkbk_pathname
print

#
# Compute storage charges.
#
if not skip_storage:
    folder_usage_dict = read_storage_usage_file(storage_usage_file)

    # Write storage usage data to BillingDetails file.
    write_storage_usage_data(folder_usage_dict, sheet_name_to_sheet_map['Storage'])

#
# Compute computing charges.
#
if not skip_computing:
    compute_computing_charges(billing_config_wkbk, begin_month_timestamp, end_month_timestamp, accounting_file,
                              sheet_name_to_sheet_map['Computing'],
                              sheet_name_to_sheet_map['Nonbillable Jobs'], sheet_name_to_sheet_map['Failed Jobs'])

#
# Compute consulting charges.
#
if not skip_consulting:
     compute_consulting_charges(billing_config_wkbk, begin_month_timestamp, end_month_timestamp, consulting_timesheet,
                                sheet_name_to_sheet_map['Consulting'])

#
# Compute cloud charges.
#
if not skip_cloud:
    compute_cloud_charges(billing_details_wkbk, google_invoice_csv, sheet_name_to_sheet_map['Cloud'])

#
# Close the output workbook and write the .xlsx file.
#
print "Closing BillingDetails workbook."
billing_details_wkbk.close()