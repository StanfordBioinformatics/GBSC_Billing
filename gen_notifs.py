#!/usr/bin/env python

#===============================================================================
#
# gen_notifs.py - Generate billing notifications for each PI for month/year.
#
# ARGS:
#   1st: the BillingConfig spreadsheet.
#
# SWITCHES:
#   --billing_details_file: Location of the BillingDetails.xlsx file (default=look in BillingRoot/<year>/<month>)
#   --accounting_file: Location of accounting file (overrides BillingConfig.xlsx)
#   --billing_root:    Location of BillingRoot directory (overrides BillingConfig.xlsx)
#                      [default if no BillingRoot in BillingConfig.xlsx or switch given: CWD]
#   --year:            Year of snapshot requested. [Default is this year]
#   --month:           Month of snapshot requested. [Default is last month]
#   --pi_sheets:       Put Billing sheets from PI-specific BillingNotifications workbooks in
#                        the BillingAggregate workbook (default=False).
#
# OUTPUT:
#   BillingNotification spreadsheets for each PI
#     in BillingRoot/<year>/<month>/GBSCBilling-<pi_tag>.<year>-<month>.xlsx
#   Various messages about current processing status to STDOUT.
#
# ASSUMPTIONS:
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
from collections import defaultdict
import datetime
import time
import os
import re
import sys

import xlrd
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell

# Simulate an "include billing_common.py".
SCRIPT_DIR = os.path.dirname(os.path.realpath(__file__))
execfile(os.path.join(SCRIPT_DIR, "billing_common.py"))

#=====
#
# CONSTANTS
#
#=====
# From billing_common.py
global BILLING_DETAILS_PREFIX
global BILLING_NOTIFS_SHEET_COLUMNS
global BILLING_AGGREG_SHEET_COLUMNS
global BILLING_NOTIFS_PREFIX
global CONSULTING_HOURS_FREE
global CONSULTING_TRAVEL_RATE_DISCOUNT

#=====
#
# GLOBALS
#
#=====

#
# For make_format(), a data structure to save all the dictionaries and resulting Format objects
#  which were created for a given workbook.
#
# Data Structure: dict with workbooks as keys, and values of [(property_dict, Format)*]
FORMAT_PROPS_PER_WORKBOOK = defaultdict(list)

#
# These globals are data structures read in from BillingConfig workbook.
#

# List of pi_tags.
pi_tag_list = []

# Mapping from usernames to list of [date, pi_tag].
username_to_pi_tag_dates = defaultdict(list)

# Mapping from usernames to a list of [email, full_name].
username_to_user_details = defaultdict(list)

# Mapping from pi_tags to list of [first_name, last_name, email].
pi_tag_to_names_email = defaultdict(list)

# Mapping from job_tags to list of [pi_tag, %age].
job_tag_to_pi_tag_pctages = defaultdict(list)

# Mapping from folders to list of [pi_tag, %age].
folder_to_pi_tag_pctages = defaultdict(list)

#
# These globals are data structures used to write the BillingNotification workbooks.
#

# Mapping from pi_tag to list of [folder, size, %age].
pi_tag_to_folder_sizes = defaultdict(list)

# Mapping from pi_tag to list of [username, cpu_core_hrs, %age].
pi_tag_to_username_cpus = defaultdict(list)

# Mapping from pi_tag to list of [job_tag, cpu_core_hrs, %age].
pi_tag_to_job_tag_cpus = defaultdict(list)

# Mapping from pi_tag to list of [job_tag, list of [username, cpu_core_hrs], %age].
pi_tag_to_job_tag_username_cpus = defaultdict(list)

# Mapping from pi_tag to list of [date, username, job_name, account, cpu_core_hrs, jobID, %age].
pi_tag_to_job_details = defaultdict(list)

# Mapping from pi_tag to list of [username, date_added, date_removed, %age].
pi_tag_to_user_details = defaultdict(list)

# Mapping from pi_tag to list of [storage_charge, computing_charge, cloud_charge, consulting_charge, total_charge].
pi_tag_to_charges = defaultdict(list)

# Mapping from pi_tag to set of (cloud account, %age) tuples.
pi_tag_to_cloud_account_pctages = defaultdict(set)

# Mapping from cloud account to set of cloud project IDs (several per project possible in this set).
cloud_account_to_cloud_projects = defaultdict(set)

# Mapping from (cloud project ID, cloud account) to lists of (platform, account, description, dates, quantity, UOM, charge) tuples.
cloud_project_account_to_cloud_details = defaultdict(list)

# Mapping from (cloud project ID, cloud account) to total charge.
cloud_project_account_to_total_charges = defaultdict(float)

# Mapping from cloud project number to cloud project ID (1-to-1 mapping).
cloud_projnum_to_cloud_project = dict()

## Bioinformatics Consulting:

# Mapping from pi_tag to list of [date, summary, notes, consultant, hours, billable_hours].
pi_tag_to_consulting_details = defaultdict(list)

# Mapping from pi_tag to list of (date, summary, consultant, hours, billable_hours].
pi_tag_to_consulting_charges = defaultdict(list)

#=====
#
# FUNCTIONS
#
#=====
# From billing_common.py
global from_timestamp_to_excel_date
global from_excel_date_to_timestamp
global from_timestamp_to_date_string
global from_excel_date_to_date_string
global from_ymd_date_to_timestamp
global sheet_get_named_column
global read_config_sheet
global config_sheet_get_dict
global filter_by_dates


# This function takes an arbitrary number of dicts with
# xlsxwriter Format properties in them, adds the format to the given workbook,
# and returns it.
#
# This function caches the ones it creates per workbook, so if a format is requested more than once,
#  it will simply return the previously created Format and not make a new one.
#
def make_format(wkbk, *prop_dicts):

    # Define the final property dict.
    final_prop_dict = dict()
    # Combine all the input dicts into the final dict.
    map(lambda d: final_prop_dict.update(d), prop_dicts)

    # Get the list of (prop_dict, Format)s for this workbook.
    prop_dict_format_list = FORMAT_PROPS_PER_WORKBOOK.setdefault(wkbk, [])

    for (prop_dict, wkbk_format) in prop_dict_format_list:
        # Is final_prop_dict already in the list?
        if final_prop_dict == prop_dict:
            # Yes: return the associated Format object.
            format_obj = wkbk_format
            break
    else:
        # Nope: new prop_dict, therefore we must make a new Format object.
        format_obj = wkbk.add_format(final_prop_dict)
        # Save the prop_dict and Format object for later use.
        prop_dict_format_list.append((final_prop_dict, format_obj))

    return format_obj


# This function creates some formats in a BillingNotification workbook,
# creates the necessary sheets, and writes the column headers in the sheets.
# It also makes the Billing sheet the active sheet when it is opened in Excel.
def init_billing_notifs_wkbk(wkbk):

    global BOLD_FORMAT
    global DATE_FORMAT
    global FLOAT_FORMAT
    global INT_FORMAT
    global MONEY_FORMAT
    global PERCENT_FORMAT

    # Create formats for use within the workbook.
    BOLD_FORMAT    = make_format(wkbk, {'bold' : True})
    DATE_FORMAT    = make_format(wkbk, {'num_format' : 'mm/dd/yy'})
    INT_FORMAT     = make_format(wkbk, {'num_format' : '0'})
    FLOAT_FORMAT   = make_format(wkbk, {'num_format' : '0.0'})
    MONEY_FORMAT   = make_format(wkbk, {'num_format' : '$#,##0.00'})
    PERCENT_FORMAT = make_format(wkbk, {'num_format' : '0%'})

    sheet_name_to_sheet = dict()

    for sheet_name in BILLING_NOTIFS_SHEET_COLUMNS:

        sheet = wkbk.add_worksheet(sheet_name)
        for col in range(0, len(BILLING_NOTIFS_SHEET_COLUMNS[sheet_name])):
            sheet.write(0, col, BILLING_NOTIFS_SHEET_COLUMNS[sheet_name][col], BOLD_FORMAT)

        sheet_name_to_sheet[sheet_name] = sheet

    # Make the Billing sheet the active one.
    sheet_name_to_sheet['Billing'].activate()

    return sheet_name_to_sheet


# This function creates a bold format in a BillingAggregate workbook,
# creates the necessary sheets, and writes the column headers in the sheets.
# It also makes the Totals sheet the active sheet when it is opened in Excel.
def init_billing_aggreg_wkbk(wkbk, pi_tag_list):

    bold_format = make_format(wkbk, {'bold' : True})

    sheet_name_to_sheet = dict()

    for sheet_name in BILLING_AGGREG_SHEET_COLUMNS:

        sheet = wkbk.add_worksheet(sheet_name)
        for col in range(0, len(BILLING_AGGREG_SHEET_COLUMNS[sheet_name])):
            sheet.write(0, col, BILLING_AGGREG_SHEET_COLUMNS[sheet_name][col], bold_format)

        sheet_name_to_sheet[sheet_name] = sheet

    if args.pi_sheets:
        # Make a sheet for each PI.
        for pi_tag in sorted(pi_tag_list):

            sheet = wkbk.add_worksheet(pi_tag)
            sheet_name_to_sheet[pi_tag] = sheet

    # Make the Aggregate sheet the active one.
    sheet_name_to_sheet['Totals'].activate()

    return sheet_name_to_sheet


# This function scans the username_to_pi_tag_dates dict to create a list of [pi_tag, %age]s
# for the PIs that the given user was working for on the given date.
def get_pi_tags_for_username_by_date(username, date_timestamp):

    # Add PI Tag to the list if the given date is after date_added, but before date_removed.

    pi_tag_list = []

    pi_tag_dates = username_to_pi_tag_dates.get(username)
    if pi_tag_dates is not None:

        date_excel = from_timestamp_to_excel_date(date_timestamp)

        for (pi_tag, date_added, date_removed, pctage) in pi_tag_dates:
            if date_added <= date_excel < date_removed:
                pi_tag_list.append([pi_tag, pctage])

    return pi_tag_list


# Creates all the data structures used to write the BillingNotification workbook.
# The overall goal is to mimic the tables of the notification sheets so that
# to build the table, all that is needed is to print out one of these data structures.
def build_global_data(wkbk, begin_month_timestamp, end_month_timestamp):

    pis_sheet      = wkbk.sheet_by_name("PIs")
    folders_sheet  = wkbk.sheet_by_name("Folders")
    users_sheet    = wkbk.sheet_by_name("Users")
    job_tags_sheet = wkbk.sheet_by_name("JobTags")

    begin_month_exceldate = from_timestamp_to_excel_date(begin_month_timestamp)
    end_month_exceldate   = from_timestamp_to_excel_date(end_month_timestamp)

    #
    # Create list of pi_tags.
    #
    global pi_tag_list

    pi_tag_list = sheet_get_named_column(pis_sheet, "PI Tag")

    #
    # Create mapping from pi_tag to a list of PI name and email.
    #
    global pi_tag_to_names_email

    pi_first_names = sheet_get_named_column(pis_sheet, "PI First Name")
    pi_last_names  = sheet_get_named_column(pis_sheet, "PI Last Name")
    pi_emails      = sheet_get_named_column(pis_sheet, "PI Email")

    pi_details_list = zip(pi_first_names, pi_last_names, pi_emails)

    pi_tag_to_names_email = dict(zip(pi_tag_list, pi_details_list))

    #
    # Organize data from the Cloud sheet, if present.
    #
    cloud_sheet = wkbk.sheet_by_name("Cloud")

    #
    # Create mapping from pi_tag to cloud project from the BillingConfig PIs sheet.
    # Create mapping from cloud project to list of (pi_tag, %age) tuples.
    # Create mapping from cloud project to cloud account (1-to-1).
    # Create mapping from cloud project to cloud project number (1-to-1).
    #
    global pi_tag_to_cloud_account_pctages
    global cloud_account_to_cloud_projects

    cloud_pi_tags     = sheet_get_named_column(cloud_sheet, "PI Tag")
    cloud_projects    = sheet_get_named_column(cloud_sheet, "Project")
    cloud_projnums    = sheet_get_named_column(cloud_sheet, "Project Number")
    cloud_projids     = sheet_get_named_column(cloud_sheet, "Project ID")
    cloud_accounts    = sheet_get_named_column(cloud_sheet, "Account")
    cloud_pctage      = sheet_get_named_column(cloud_sheet, "%age")

    cloud_dates_added = sheet_get_named_column(cloud_sheet, "Date Added")
    cloud_dates_remvd = sheet_get_named_column(cloud_sheet, "Date Removed")

    cloud_rows = filter_by_dates(zip(cloud_pi_tags, cloud_projects, cloud_projnums, cloud_projids,
                                     cloud_accounts, cloud_pctage),
                                 zip(cloud_dates_added, cloud_dates_remvd),
                                 begin_month_exceldate, end_month_exceldate)

    for (pi_tag, project, projnum, projid, account, pctage) in cloud_rows:

        # Associate the project name and percentage to be charged with the pi_tag.
        pi_tag_to_cloud_account_pctages[pi_tag].add((account, pctage))

        # Associate the project number with the pi_tag also, in case the project is deleted and loses its name.
        #pi_tag_to_cloud_account_pctages[pi_tag].add((projnum, pctage))

        # Associate the account with the project name, the project number, and the project ID.
        cloud_account_to_cloud_projects[account].add(project)
        cloud_account_to_cloud_projects[account].add(projnum)
        cloud_account_to_cloud_projects[account].add(projid)
        cloud_account_to_cloud_projects[account].add("")  # For credits to the account.

        # Associate the project ID with its project number.
        cloud_projnum_to_cloud_project[projnum] = projid

    #
    # Filter pi_tag_list for PIs active in the current month.
    #
    pi_dates_added   = sheet_get_named_column(pis_sheet, "Date Added")
    pi_dates_removed = sheet_get_named_column(pis_sheet, "Date Removed")

    pi_tags_and_dates_added = zip(pi_tag_list, pi_dates_added, pi_dates_removed)

    for (pi_tag, date_added, date_removed) in pi_tags_and_dates_added:

        # Convert the Excel dates to timestamps.
        date_added_timestamp = from_excel_date_to_timestamp(date_added)
        if date_removed != '':
            date_removed_timestamp = from_excel_date_to_timestamp(date_removed)
        else:
            date_removed_timestamp = end_month_timestamp + 1  # Not in this month.

        # If the date added is AFTER the end of this month, or
        #  the date removed is BEFORE the beginning of this month,
        # then remove the pi_tag from the list.
        if date_added_timestamp >= end_month_timestamp:

            print >> sys.stderr, " *** Ignoring PI %s: added after this month on %s" % (pi_tag_to_names_email[pi_tag][1], from_excel_date_to_date_string(date_added))
            pi_tag_list.remove(pi_tag)

        elif date_removed_timestamp < begin_month_timestamp:

            print >> sys.stderr, " *** Ignoring PI %s: removed before this month on %s" % (pi_tag_to_names_email[pi_tag][1], from_excel_date_to_date_string(date_removed))
            pi_tag_list.remove(pi_tag)

    #
    # Create mapping from usernames to a list of user details.
    #
    global username_to_user_details

    usernames  = sheet_get_named_column(users_sheet, "Username")
    emails     = sheet_get_named_column(users_sheet, "Email")
    full_names = sheet_get_named_column(users_sheet, "Full Name")

    username_details_rows = zip(usernames, emails, full_names)

    for (username, email, full_name) in username_details_rows:
        username_to_user_details[username] = [email, full_name]

    #
    # Create mapping from usernames to a list of pi_tag/dates.
    #
    global username_to_pi_tag_dates

    pi_tags       = sheet_get_named_column(users_sheet, "PI Tag")
    dates_added   = sheet_get_named_column(users_sheet, "Date Added")
    dates_removed = sheet_get_named_column(users_sheet, "Date Removed")
    pctages       = sheet_get_named_column(users_sheet, "%age")

    username_rows = zip(usernames, pi_tags, dates_added, dates_removed, pctages)

    for (username, pi_tag, date_added, date_removed, pctage) in username_rows:
        username_to_pi_tag_dates[username].append([pi_tag, date_added, date_removed, pctage])

    #
    # Create mapping from pi_tags to a list of [username, date_added, date_removed]
    #
    global pi_tag_to_user_details

    for username in username_to_pi_tag_dates:

        pi_tag_date_list = username_to_pi_tag_dates[username]

        for (pi_tag, date_added, date_removed, pctage) in pi_tag_date_list:
            pi_tag_to_user_details[pi_tag].append([username, date_added, date_removed, pctage])

    #
    # Create mapping from job_tag to list of pi_tags and %ages.
    #
    global job_tag_to_pi_tag_pctages

    job_tags = sheet_get_named_column(job_tags_sheet, "Job Tag")
    pi_tags  = sheet_get_named_column(job_tags_sheet, "PI Tag")
    pctages  = sheet_get_named_column(job_tags_sheet, "%age")

    dates_added   = sheet_get_named_column(job_tags_sheet, "Date Added")
    dates_removed = sheet_get_named_column(job_tags_sheet, "Date Removed")

    job_tag_rows = filter_by_dates(zip(job_tags, pi_tags, pctages), zip(dates_added, dates_removed),
                                   begin_month_exceldate, end_month_exceldate)

    for (job_tag, pi_tag, pctage) in job_tag_rows:
        job_tag_to_pi_tag_pctages[job_tag].append([pi_tag, pctage])

    # Add pi_tags and "baas_pi_tag" as job_tags.
    for pi_tag in pi_tag_list:
        job_tag_to_pi_tag_pctages[pi_tag].append([pi_tag, 1.0])
        job_tag_to_pi_tag_pctages["baas_" + pi_tag].append([pi_tag, 1.0])

    #
    # Create mapping from folder to list of pi_tags and %ages.
    #
    global folder_to_pi_tag_pctages

    # Get the Folders from PI Sheet
    folders = sheet_get_named_column(pis_sheet, "PI Folder")
    pi_tags = sheet_get_named_column(pis_sheet, "PI Tag")
    pctages = [1.0] * len(folders)

    dates_added   = sheet_get_named_column(pis_sheet, "Date Added")
    dates_removed = sheet_get_named_column(pis_sheet, "Date Removed")

    # Add the Folders from Folder sheet
    folders += sheet_get_named_column(folders_sheet, "Folder")
    pi_tags += sheet_get_named_column(folders_sheet, "PI Tag")
    pctages += sheet_get_named_column(folders_sheet, "%age")

    dates_added   += sheet_get_named_column(folders_sheet, "Date Added")
    dates_removed += sheet_get_named_column(folders_sheet, "Date Removed")

    folder_rows = filter_by_dates(zip(folders, pi_tags, pctages), zip(dates_added, dates_removed),
                                  begin_month_exceldate, end_month_exceldate)

    for (folder, pi_tag, pctage) in folder_rows:

        # Account for multiple folders separated by commas.
        pi_folder_list = folder.split(',')

        for pi_folder in pi_folder_list:
            folder_to_pi_tag_pctages[pi_folder].append([pi_tag, pctage])


# Reads the particular rate requested from the Rates sheet of the BillingConfig workbook.
def get_rates(wkbk, rate_type):

    rates_sheet = wkbk.sheet_by_name('Rates')

    types   = sheet_get_named_column(rates_sheet, 'Type')
    amounts = sheet_get_named_column(rates_sheet, 'Amount')

    for (type, amount) in zip(types, amounts):
        if type == rate_type:
            return amount
    else:
        return None

def get_rate_a1_cell(wkbk, rate_type):

    rates_sheet = wkbk.sheet_by_name('Rates')

    header_row = rates_sheet.row_values(0)

    # Find the column numbers for 'Type' and 'Amount'.
    type_col = -1
    amt_col = -1
    for idx in range(len(header_row)):
        if header_row[idx] == 'Type':
            type_col = idx
        elif header_row[idx] == 'Amount':
            amt_col = idx

    if type_col == -1 or amt_col == -1:
        return None

    # Get column of 'Types'.
    types = rates_sheet.col_values(type_col, start_rowx=1)

    # When you find the row with rate_type, return the Amount col and this row.
    for idx in range(len(types)):
        if types[idx] == rate_type:
            # +1 is for "GBSC Rates:" above header line, +1 is for header line.
            return 'Rates!%s' % xl_rowcol_to_cell(idx + 1 + 1, amt_col, True, True)
    else:
        return None


# Reads the Storage sheet of the BillingDetails workbook given, and populates
# the pi_tag_to_folder_sizes dict with the folder measurements for each PI.
def read_storage_sheet(wkbk):

    global pi_tag_to_folder_sizes

    storage_sheet = wkbk.sheet_by_name("Storage")

    for row in range(1,storage_sheet.nrows):

        (date, timestamp, folder, size, used) = storage_sheet.row_values(row)

        # List of [pi_tag, %age] pairs.
        pi_tag_pctages = folder_to_pi_tag_pctages[folder]

        for (pi_tag, pctage) in pi_tag_pctages:
            pi_tag_to_folder_sizes[pi_tag].append([folder, size, pctage])


# Reads the Computing sheet of the BillingDetails workbook given, and populates
# the job_tag_to_pi_tag_cpus, pi_tag_to_job_tag_cpus, pi_tag_to_username_cpus, and
# pi_tag_to_job_details dicts.
def read_computing_sheet(wkbk):

    global pi_tag_to_job_details
    global pi_tag_to_job_tag_cpus
    global pi_tag_to_username_cpus

    computing_sheet = wkbk.sheet_by_name("Computing")

    for row in range(1, computing_sheet.nrows):

        (job_date, job_timestamp, job_username, job_name, account, node, cores, wallclock, jobID) = \
            computing_sheet.row_values(row)

        # Calculate CPU-core-hrs for job.
        cpu_core_hrs = cores * wallclock / 3600.0  # wallclock is in seconds.

        # Rename this variable for easier understanding.
        job_tag = account.lower()

        if job_tag != '':
            job_pi_tag_pctage_list = job_tag_to_pi_tag_pctages[job_tag]
        else:
            # No job tag means credit the job to the user's lab.
            job_pi_tag_pctage_list = get_pi_tags_for_username_by_date(job_username, job_timestamp)

        if len(job_pi_tag_pctage_list) == 0:
            print "   *** No PI associated with job ID %d" % jobID
            continue

        # Distribute this job's CPU-hrs amongst pi_tags by %ages.
        for (pi_tag, pctage) in job_pi_tag_pctage_list:

            # This list is [job_tag, list of [username, cpu_core_hrs, %age], %age].
            job_tag_username_cpu_list = pi_tag_to_job_tag_username_cpus.get(pi_tag)

            # If pi_tag has an existing list of job_tag/username/CPUs:
            if job_tag_username_cpu_list is not None:

                # Find if job tag for job is in list of job_tag/CPUs for this pi_tag.
                for (pi_job_tag, pi_username_cpu_list, pi_pctage) in job_tag_username_cpu_list:

                    # If the job tag we are looking at is the one from the present job:
                    if pi_job_tag == job_tag and pi_pctage == pctage:

                        # Find job username in list for job tag:
                        for username_cpu in pi_username_cpu_list:
                            if job_username == username_cpu[0]:
                                username_cpu[1] += cpu_core_hrs
                                break
                        else:
                            pi_username_cpu_list.append([job_username, cpu_core_hrs])

                        # Leave job_tag_username_cpu_list loop.
                        break

                else:
                    # No matching job_tag in pi_tag list -- add a new one to the list.
                    job_tag_username_cpu_list.append([job_tag, [[job_username, cpu_core_hrs]], pctage])

            # Else start a new job_tag/CPUs list for the pi_tag.
            else:
                pi_tag_to_job_tag_username_cpus[pi_tag] = [[job_tag, [[job_username, cpu_core_hrs]], pctage]]

            #
            # Save job details for pi_tag.
            #
            new_job_details = [job_date, job_username, job_name, account, node, cpu_core_hrs, jobID, pctage]
            pi_tag_to_job_details[pi_tag].append(new_job_details)



# Read the Cloud sheet from the BillingDetails workbook.
def read_cloud_sheet(wkbk):

    cloud_sheet = wkbk.sheet_by_name('Cloud')

    for row in range(1,cloud_sheet.nrows):

        (platform, account, project, description, dates, quantity, uom, charge) = cloud_sheet.row_values(row)

        # If project is of the form "<project name>(<project-id>)" or "<project name>[<project-id>]", get the "<project-id>".
        project_re = re.search("[(\[]([a-z0-9-:.]+)[\])]", project)
        if project_re is not None:
            project = project_re.group(1)
        else:
            pass  # If no parens, use the original project name.

        # Save the cloud item in a list of charges for that PI.
        cloud_project_account_to_cloud_details[(project, account)].append((platform, description, dates, quantity, uom, charge))

        # Accumulate the total cost of a project.
        cloud_project_account_to_total_charges[(project, account)] += float(charge)


# Reads the Consulting sheet of the BillingDetails workbook.
def read_consulting_sheet(wkbk):

    consulting_sheet = wkbk.sheet_by_name('Consulting')

    dates       = sheet_get_named_column(consulting_sheet, 'Date')
    pi_tags     = sheet_get_named_column(consulting_sheet, 'PI Tag')
    hours       = sheet_get_named_column(consulting_sheet, 'Hours')
    travel_hours= sheet_get_named_column(consulting_sheet, 'Travel Hours')
    consultants = sheet_get_named_column(consulting_sheet, 'Participants')
    summaries   = sheet_get_named_column(consulting_sheet, 'Summary')
    notes       = sheet_get_named_column(consulting_sheet, 'Notes')
    cumul_hours = sheet_get_named_column(consulting_sheet, 'Cumul Hours')

    consulting_details = zip(dates, pi_tags, hours, travel_hours, consultants, summaries, notes, cumul_hours)

    for (date, pi_tag, hours, travel_hours, consultant, summary, notes, cumul_hours) in consulting_details:

        # Save the consulting item in a list of details for each PI.
        pi_tag_to_consulting_details[pi_tag].append((date, summary, notes, consultant, float(hours), float(travel_hours), float(cumul_hours)))

        #
        # Calculate the number of free hours and billable hours in this transaction.
        #
        start_hours_used = float(cumul_hours) - float(hours) - float(travel_hours)

        if start_hours_used < CONSULTING_HOURS_FREE:
            free_hours_remaining = CONSULTING_HOURS_FREE - start_hours_used
        else:
            free_hours_remaining = 0

        if hours < free_hours_remaining:
            free_hours_used = hours
        else:
            free_hours_used = free_hours_remaining

        billable_hours = hours - free_hours_used + (travel_hours * CONSULTING_TRAVEL_RATE_DISCOUNT)

        # Save the consulting charges in a list of items for each PI.
        pi_tag_to_consulting_charges[pi_tag].append((date, summary, consultant, float(hours), float(travel_hours), float(billable_hours)))


# Generates the Billing sheet of a BillingNotifications (or BillingAggregate) workbook for a particular pi_tag.
# It uses dicts pi_tag_to_folder_sizes, pi_tag_to_username_cpus, and pi_tag_to_job_tag_cpus, and puts
# summaries of its results in dict pi_tag_to_charges.
def generate_billing_sheet(wkbk, sheet, pi_tag, begin_month_timestamp, end_month_timestamp):

    global pi_tag_to_charges

    #
    # Set the column and row widths to contain all our data.
    #

    # Give the first column 1 unit of space.
    sheet.set_column('A:A', 1)
    # Give the second column 35 units of space.
    sheet.set_column('B:B', 40)
    # Give the third, fourth, and fifth columns 10 units of space each.
    sheet.set_column('C:C', 11)
    sheet.set_column('D:D', 11)
    sheet.set_column('E:E', 11)
    # Give the first row 50 units of space.  ("Bill for Services Rendered")
    sheet.set_row(0, 50)
    # Give the second row 30 units of space. ("PI: <PI NAME>")
    sheet.set_row(1, 30)

    #
    # Write out the Document Header first ("Bill for Services Rendered")
    #

    # Write the text of the first row, with the GBSC address in merged columns.
    fmt = make_format(wkbk, {'font_size': 18, 'bold': True, 'underline': True,
                             'align': 'left', 'valign': 'vcenter'})
    sheet.write(0, 1, 'Bill for Services Rendered', fmt)

    fmt = make_format(wkbk, {'font_size': 12, 'text_wrap': True})
    sheet.merge_range('C1:F1', "Genetics Bioinformatics Service Center (GBSC)\nSoM Technology & Innovation Center\n3165 Porter Drive, Palo Alto, CA", fmt)

    # Write the PI name on the second row.
    (pi_first_name, pi_last_name, _) = pi_tag_to_names_email[pi_tag]

    fmt = make_format(wkbk, {'font_size' : 16, 'align': 'left', 'valign': 'vcenter'})
    sheet.write(1, 1, "PI: %s, %s" % (pi_last_name, pi_first_name), fmt)

    #
    # Write the Billing Period dates on the fourth row.
    #
    begin_date_string = from_timestamp_to_date_string(begin_month_timestamp)

    # If we are running this script mid-month, use today's date as the end date for the Billing Period.
    now_timestamp = time.time()
    if now_timestamp < end_month_timestamp:
        end_date_string = from_timestamp_to_date_string(now_timestamp)
    else:
        end_date_string = from_timestamp_to_date_string(end_month_timestamp-1)

    billing_period_string = "Billing Period: %s - %s" % (begin_date_string, end_date_string)

    fmt = make_format(wkbk, { 'font_size': 14, 'bold': True})
    sheet.write(3, 1, billing_period_string, fmt)

    #
    # Calculate Breakdown of Charges first, then use those cumulative
    #  totals to fill out the Summary of Charges.
    #

    # Set up some formats for use in these tables.
    border_style = 1

    # For "Summary of Charges" and "Breakdown of Charges"
    top_header_fmt = make_format(wkbk, {'font_size': 14, 'bold': True, 'underline': True,
                                        'align': 'right',
                                        'top': border_style, 'left': border_style})
    # For "Storage", "Computing", "Cloud Services", and "Bioinformatics Consulting" headers
    header_fmt = make_format(wkbk, {'font_size': 12, 'bold': True, 'underline': True,
                                    'align': 'left',
                                    'left': border_style})
    # Same as above, but with no underline: for the Summary table.
    header_no_ul_fmt = make_format(wkbk, {'font_size': 12, 'bold': True,
                                          'align': 'right',
                                          'left': border_style})

    # For subheaders within subtables, like "Account: XXX" in Computing subtable.
    sub_header_fmt = make_format(wkbk, {'font_size': 11, 'bold': True,
                                        'align': 'right', 'underline': True,
                                        'left': border_style})

    # For column headers in subtables.
    col_header_fmt = make_format(wkbk, {'font_size': 11, 'bold': True,
                                        'align': 'right'})
    # As above, but allowing text wrap for long column headers (see Bioinformatics Computing subtable)
    col_header_textwrap_fmt = make_format(wkbk, {'font_size': 11, 'bold': True,
                                                 'align': 'right',
                                                 'text_wrap': True})
    # As sub_header_fmt, but with a border on the left.
    col_header_left_fmt = make_format(wkbk, {'font_size': 11, 'bold': True,
                                             'align': 'right',
                                             'left': border_style})
    # As sub_header_fmt, but with a border on the right.
    col_header_right_fmt = make_format(wkbk, {'font_size': 11, 'bold': True,
                                              'align': 'right',
                                              'right': border_style})

    # Text entry in a subtable (border on the left).
    item_entry_fmt = make_format(wkbk, {'font_size': 10,
                                        'align': 'right',
                                        'left': border_style})
    # As above, plus allowing text wrap (for descriptions in Bioinformatics Consulting).
    item_entry_textwrap_fmt = make_format(wkbk, {'font_size': 10,
                                                 'align': 'right',
                                                 'left': border_style,
                                                 'text_wrap': True})
    item_entry_italics_fmt = make_format(wkbk, {'font_size': 10,
                                                'align': 'right',
                                                'left': border_style,
                                                'italics': True})
    # Float entry in a subtable.
    float_entry_fmt = make_format(wkbk, {'font_size': 10,
                                         'align': 'right',
                                         'num_format': '0.0'})
    # As above, but vertically aligned to top (for Bioinformatics Consulting table).
    float_entry_valign_top_fmt = make_format(wkbk, {'font_size': 10,
                                                     'align': 'right',
                                                     'valign': 'top',
                                                     'num_format': '0.0'})
    # Integer entry in a subtable (not used?).
    int_entry_fmt = make_format(wkbk, {'font_size': 10,
                                       'align': 'right',
                                       'num_format': '0'})
    # Percentage entry in a subtable.
    pctage_entry_fmt = make_format(wkbk, {'font_size': 10,
                                          'num_format': '0%'})
    # String entry in a subtable, aligned right.
    string_entry_fmt = make_format(wkbk,{'font_size': 10,
                                         'align': 'right'})
    # As above, but vertically aligned to top (for Bioinformatics Consulting table).
    string_entry_valign_top_fmt = make_format(wkbk,{'font_size': 10,
                                                    'align': 'right',
                                                    'valign': 'top'})
    # Cost entry in Cloud Services subtable.
    cost_fmt = make_format(wkbk, {'font_size': 10,
                                  'align': 'right',
                                  'num_format': '$#,##0.00'})

    # Charge entry in subtables (with border on right).
    charge_fmt = make_format(wkbk, {'font_size': 10,
                                    'align': 'right',
                                    'right': border_style,
                                    'num_format': '$#,##0.00'})
    # As above, only vertically aligned to the top.
    charge_valign_top_fmt = make_format(wkbk, {'font_size': 10,
                                               'align': 'right',
                                               'valign': 'top',
                                               'right': border_style,
                                               'num_format': '$#,##0.00'})
    # Charge entry in Summary of Charges table.
    big_charge_fmt = make_format(wkbk, {'font_size': 12,
                                        'align': 'right',
                                        'right': border_style,
                                        'num_format': '$#,##0.00'})
    # As above, only bold (for grand total).
    big_bold_charge_fmt = make_format(wkbk, {'font_size': 12, 'bold': True,
                                             'align': 'right',
                                             'right': border_style, 'bottom': border_style,
                                             'num_format': '$#,##0.00'})
    # "Total XXX" entry for subtotals within subtables.
    bot_header_fmt = make_format(wkbk, {'font_size': 14, 'bold': True,
                                        'align': 'right',
                                        'left': border_style})
    # As above, but with a bottom border.
    bot_header_border_fmt = make_format(wkbk, {'font_size': 14, 'bold': True,
                                               'align': 'right',
                                               'left': border_style,
                                               'bottom': border_style})
    # Formats for borders in cells.
    upper_right_border_fmt = make_format(wkbk, {'top': border_style, 'right': border_style})
    lower_right_border_fmt = make_format(wkbk, {'bottom': border_style, 'right': border_style})
    lower_left_border_fmt  = make_format(wkbk, {'bottom': border_style, 'left': border_style})
    left_border_fmt = make_format(wkbk, {'left': border_style})
    right_border_fmt = make_format(wkbk, {'right': border_style})
    top_border_fmt = make_format(wkbk, {'top': border_style})
    bottom_border_fmt = make_format(wkbk, {'bottom': border_style})

    ######
    #
    # "Breakdown of Charges" (B14:??)
    #
    ######

    # Start the Breakdown of Charges table on the fifteenth row.
    curr_row = 14
    sheet.write(curr_row, 1, "Breakdown of Charges:", top_header_fmt)
    sheet.write(curr_row, 2, None, top_border_fmt)
    sheet.write(curr_row, 3, None, top_border_fmt)
    sheet.write(curr_row, 4, None, upper_right_border_fmt)
    curr_row += 1

    ###
    #
    # STORAGE Subtable of Breakdown of Charges table
    #
    ###

    # Skip line between "Breakdown of Charges".
    sheet.write(curr_row, 1, None, left_border_fmt)
    sheet.write(curr_row, 4, None, right_border_fmt)
    curr_row += 1
    # Write the "Storage" line.
    sheet.write(curr_row, 1, "Storage", header_fmt)
    sheet.write(curr_row, 4, None, right_border_fmt)
    curr_row += 1
    # Write the storage headers.
    sheet.write(curr_row, 1, "Folder", col_header_left_fmt)
    sheet.write(curr_row, 2, "Storage (Tb)", col_header_fmt)
    sheet.write(curr_row, 3, "%age", col_header_fmt)
    sheet.write(curr_row, 4, "Charge", col_header_right_fmt)
    curr_row += 1

    total_storage_charges = 0.0
    total_storage_sizes   = 0.0

    # Get the rate from the Rates sheet of the BillingConfig workbook.
    rate_tb_per_month = get_rates(billing_config_wkbk, 'Local HPC Storage')

    starting_cloud_row = curr_row
    ending_cloud_row   = curr_row
    for (folder, size, pctage) in pi_tag_to_folder_sizes[pi_tag]:
        sheet.write(curr_row, 1, folder, item_entry_fmt)
        sheet.write(curr_row, 2, size, float_entry_fmt)
        sheet.write(curr_row, 3, pctage, pctage_entry_fmt)

        # Calculate charges.
        if rate_tb_per_month is not None:
            charge = size * pctage * rate_tb_per_month
            total_storage_charges += charge
        else:
            charge = "No rate"

        total_storage_sizes += size

        cost_a1_cell   = xl_rowcol_to_cell(curr_row, 2)
        pctage_a1_cell = xl_rowcol_to_cell(curr_row, 3)
        sheet.write_formula(curr_row, 4, '=%s*%s*%s' % (cost_a1_cell, pctage_a1_cell,
                                                        get_rate_a1_cell(billing_config_wkbk, 'Local HPC Storage')),
                            charge_fmt, charge)

        # Keep track of last row with storage values.
        ending_cloud_row = curr_row

        # Advance to the next row.
        curr_row += 1

    # Skip the line before Total Storage.
    sheet.write(curr_row, 1, None, left_border_fmt)
    sheet.write(curr_row, 4, None, right_border_fmt)
    curr_row += 1

    # Write the Total Storage line.
    sheet.write(curr_row, 1, "Total Storage", bot_header_fmt)
    # sheet.write(curr_row, 2, total_storage_sizes, float_entry_fmt)
    top_sizes_a1_cell = xl_rowcol_to_cell(starting_cloud_row, 2)
    bot_sizes_a1_cell = xl_rowcol_to_cell(ending_cloud_row, 2)
    sheet.write_formula(curr_row, 2, '=SUM(%s:%s)' % (top_sizes_a1_cell, bot_sizes_a1_cell),
                        float_entry_fmt, total_storage_sizes)
    # sheet.write(curr_row, 4, total_storage_charges, charge_fmt)
    top_billable_hours_a1_cell = xl_rowcol_to_cell(starting_cloud_row, 4)
    bot_hours_a1_cell = xl_rowcol_to_cell(ending_cloud_row, 4)
    sheet.write_formula(curr_row, 4, '=SUM(%s:%s)' % (top_billable_hours_a1_cell, bot_hours_a1_cell),
                        charge_fmt, total_storage_charges)

    # Save reference to this cell for use in Summary subtable.
    total_storage_charges_a1_cell = xl_rowcol_to_cell(curr_row, 4)

    curr_row += 1

    # Skip the next line and draw line under this row.
    sheet.write(curr_row, 1, None, lower_left_border_fmt)
    sheet.write(curr_row, 2, None, bottom_border_fmt)
    sheet.write(curr_row, 3, None, bottom_border_fmt)
    sheet.write(curr_row, 4, None, lower_right_border_fmt)
    curr_row += 1

    ###
    #
    # COMPUTING Subtable of Breakdown of Charges table
    #
    ###

    # Get the rate from the Rates sheet of the BillingConfig workbook.
    rate_cpu_per_hour = get_rates(billing_config_wkbk, 'Local Computing')

    # Skip row before Computing header.
    sheet.write(curr_row, 1, None, left_border_fmt)
    sheet.write(curr_row, 4, None, right_border_fmt)
    curr_row += 1
    # Write the Computing line.
    sheet.write(curr_row, 1, "Computing", header_fmt)
    sheet.write(curr_row, 4, None, right_border_fmt)
    curr_row += 1

    # Loop over pi_tag_to_job_tag_username_cpus for job tag/username combos.
    job_tag_username_cpus_list = pi_tag_to_job_tag_username_cpus.get(pi_tag)

    # The list of "Total Charges" rows for each job tag.
    total_computing_charges_row_list = []

    if job_tag_username_cpus_list is not None:

        for (job_tag, username_cpu_list, pctage) in job_tag_username_cpus_list:

            # Write the account subheader.
            if job_tag != "":
                sheet.write(curr_row, 1, "Account: %s" % job_tag, sub_header_fmt)
            else:
                sheet.write(curr_row, 1, "Account: Lab Default", sub_header_fmt)
            sheet.write(curr_row, 4, None, col_header_right_fmt)
            curr_row += 1

            # Skip row after account subheader.
            sheet.write(curr_row, 1, None, left_border_fmt)
            sheet.write(curr_row, 4, None, right_border_fmt)
            curr_row += 1

            # Write the computing headers.
            sheet.write(curr_row, 1, "User", col_header_left_fmt)
            sheet.write(curr_row, 2, "CPU-core-hrs", col_header_fmt)
            sheet.write(curr_row, 3, "%age", col_header_fmt)
            sheet.write(curr_row, 4, "Charge", col_header_right_fmt)
            curr_row += 1

            total_computing_charges = 0.0
            total_computing_cpuhrs  = 0.0

            # Get the job details for the users associated with this PI.
            starting_computing_row = curr_row
            ending_computing_row   = curr_row
            if len(username_cpu_list) > 0:

                for (username, cpu_core_hrs) in username_cpu_list:

                    pi_tags_for_username = get_pi_tags_for_username_by_date(username, begin_month_timestamp)

                    if pi_tag in map(lambda pi_pct: pi_pct[0], pi_tags_for_username):
                        username_fmt = item_entry_fmt
                    else:
                        username_fmt = item_entry_italics_fmt

                    fullname = username_to_user_details[username][1]
                    sheet.write(curr_row, 1, "%s (%s)" % (fullname, username), username_fmt)
                    sheet.write(curr_row, 2, cpu_core_hrs, float_entry_fmt)
                    sheet.write(curr_row, 3, pctage, pctage_entry_fmt)

                    if rate_cpu_per_hour is not None:
                        charge = cpu_core_hrs * pctage * rate_cpu_per_hour
                        total_computing_charges += charge
                    else:
                        charge = "No rate"

                    # Check if user has accumulated more than $500 in a month.
                    if charge > 500:
                        print "  *** User %s (%s) for PI %s, Account %s: $%0.02f" % (username_to_user_details[username][1], username, pi_tag, job_tag, charge)

                    total_computing_cpuhrs += cpu_core_hrs

                    cpu_a1_cell    = xl_rowcol_to_cell(curr_row, 2)
                    pctage_a1_cell = xl_rowcol_to_cell(curr_row, 3)
                    sheet.write_formula(curr_row, 4, '=%s*%s*%s' % (cpu_a1_cell, pctage_a1_cell,
                                                                    get_rate_a1_cell(billing_config_wkbk, 'Local Computing')),
                                        charge_fmt, charge)

                    # Keep track of last row with computing values.
                    ending_computing_row = curr_row

                    # Advance to the next row.
                    curr_row += 1

                # Skip row after last user.
                sheet.write(curr_row, 1, None, left_border_fmt)
                sheet.write(curr_row, 4, None, right_border_fmt)
                curr_row += 1

            else:
                # No users for this PI.
                sheet.write(curr_row, 1, "No jobs", item_entry_fmt)
                sheet.write(curr_row, 4, 0.0, charge_fmt)
                curr_row += 1

            # Write the Total Charges line header.
            if job_tag != "":
                sheet.write(curr_row, 1, "Total charges: %s:" % job_tag, col_header_left_fmt)
            else:
                sheet.write(curr_row, 1, "Total charges: Lab Default:", col_header_left_fmt)

            # Write the formula for the CPU-core-hrs subtotal for the job tag.
            top_cpu_a1_cell = xl_rowcol_to_cell(starting_computing_row, 2)
            bot_cpu_a1_cell = xl_rowcol_to_cell(ending_computing_row, 2)
            sheet.write_formula(curr_row, 2, '=SUM(%s:%s)' % (top_cpu_a1_cell, bot_cpu_a1_cell),
                                float_entry_fmt, total_computing_cpuhrs)

            sheet.write(curr_row, 3, None, col_header_fmt)

            # Write the formula for the charges subtotal for the job tag.
            top_charge_a1_cell = xl_rowcol_to_cell(starting_computing_row, 4)
            bot_charge_a1_cell = xl_rowcol_to_cell(ending_computing_row, 4)
            sheet.write_formula(curr_row, 4, '=SUM(%s:%s)' % (top_charge_a1_cell, bot_charge_a1_cell),
                                charge_fmt, charge)

            # Save row of this total charges for the job tag for Total Computing charges sum.
            total_computing_charges_row_list.append(curr_row)

            curr_row += 1

            # Skip row after job tag subtotal.
            sheet.write(curr_row, 1, None, left_border_fmt)
            sheet.write(curr_row, 4, None, right_border_fmt)
            curr_row += 1

    # Write the Total Computing line.
    sheet.write(curr_row, 1, "Total Computing:", bot_header_fmt)
    # sheet.write(curr_row, 2, total_computing_cpuhrs, float_entry_fmt)
    # sheet.write(curr_row, 4, total_computing_charges, charge_fmt)

    if len(total_computing_charges_row_list) > 0:

        total_cpuhours_cell_list = map(lambda x: xl_rowcol_to_cell(x, 2), total_computing_charges_row_list)
        total_computing_charges_cell_list = map(lambda x: xl_rowcol_to_cell(x, 4), total_computing_charges_row_list)

        # Create formula from job tag total CPU-hours cells.
        total_cpuhours_formula = "=" + "+".join(total_cpuhours_cell_list)

        sheet.write_formula(curr_row, 2, total_cpuhours_formula, float_entry_fmt)

        # Create formula from job tag total charges cells.
        total_computing_charges_formula = "=" + "+".join(total_computing_charges_cell_list)

        sheet.write_formula(curr_row, 4, total_computing_charges_formula, charge_fmt)

    else:

        sheet.write(curr_row, 4, 0.0, charge_fmt)

    # Save reference to this cell for use in Summary subtable.
    total_computing_charges_a1_cell = xl_rowcol_to_cell(curr_row, 4)

    curr_row += 1

    # Skip the next line and draw line under this row.
    sheet.write(curr_row, 1, None, lower_left_border_fmt)
    sheet.write(curr_row, 2, None, bottom_border_fmt)
    sheet.write(curr_row, 3, None, bottom_border_fmt)
    sheet.write(curr_row, 4, None, lower_right_border_fmt)
    curr_row += 1

    ###
    #
    # CLOUD SERVICES Subtable of Breakdown of Charges table
    #
    ###

    # Skip line between previous subtable.
    sheet.write(curr_row, 1, None, left_border_fmt)
    sheet.write(curr_row, 4, None, right_border_fmt)
    curr_row += 1
    # Write the "Cloud Services" line.
    sheet.write(curr_row, 1, "Cloud Services", header_fmt)
    sheet.write(curr_row, 4, None, right_border_fmt)
    curr_row += 1
    # Write the cloud services headers.
    sheet.write(curr_row, 1, "Project", col_header_left_fmt)
    sheet.write(curr_row, 2, "Cost", col_header_fmt)
    sheet.write(curr_row, 3, "%age", col_header_fmt)
    sheet.write(curr_row, 4, "Charge", col_header_right_fmt)
    curr_row += 1

    total_cloud_charges = 0.0

    starting_cloud_row = curr_row
    ending_cloud_row   = curr_row - 1 # Inverted order means no projects found.
    for (account, pctage) in pi_tag_to_cloud_account_pctages[pi_tag]:

        for project in cloud_account_to_cloud_projects[account]:

            project_cost = cloud_project_account_to_total_charges[(project, account)]

            if project_cost != 0.0:
                # A blank project name means (usually) a credit applied to the account.
                if len(project) > 0:
                    # If we have the project number here, use the project name.
                    if project[0].isdigit():
                        sheet.write(curr_row, 1, cloud_projnum_to_cloud_project[project], item_entry_fmt)
                    else:
                        sheet.write(curr_row, 1, project, item_entry_fmt)
                else:
                    sheet.write(curr_row, 1, "Misc charges/credits", item_entry_fmt)
                sheet.write(curr_row, 2, project_cost, cost_fmt)
                sheet.write(curr_row, 3, pctage, pctage_entry_fmt)

                # Calculate charges.
                charge = project_cost * pctage
                total_cloud_charges += charge

                # Write formula for charges to the sheet.
                cost_a1_cell   = xl_rowcol_to_cell(curr_row, 2)
                pctage_a1_cell = xl_rowcol_to_cell(curr_row, 3)
                sheet.write_formula(curr_row, 4, '=%s*%s' % (cost_a1_cell, pctage_a1_cell),
                                    charge_fmt, charge)

                # Keep track of last row with cloud project values.
                ending_cloud_row = curr_row

                # Advance to the next row.
                curr_row += 1

    # If there were no projects, put a row saying so.
    if starting_cloud_row > ending_cloud_row:
        sheet.write(curr_row, 1, "No Projects", item_entry_fmt)
        sheet.write(curr_row, 4, 0.0, charge_fmt)
        curr_row += 1
        ending_cloud_row = starting_cloud_row

    # Skip the line before "Total Cloud Services".
    sheet.write(curr_row, 1, None, left_border_fmt)
    sheet.write(curr_row, 4, None, right_border_fmt)
    curr_row += 1

    # Write the "Total Cloud Services" line.
    sheet.write(curr_row, 1, "Total Cloud Services", bot_header_fmt)
    top_billable_hours_a1_cell = xl_rowcol_to_cell(starting_cloud_row, 4)
    bot_hours_a1_cell = xl_rowcol_to_cell(ending_cloud_row, 4)
    sheet.write_formula(curr_row, 4, '=SUM(%s:%s)' % (top_billable_hours_a1_cell, bot_hours_a1_cell),
                        charge_fmt, total_cloud_charges)

    # Save reference to this cell for use in Summary subtable.
    total_cloud_charges_a1_cell = xl_rowcol_to_cell(curr_row, 4)

    curr_row += 1

    # Skip the next line and draw line under this row.
    sheet.write(curr_row, 1, None, lower_left_border_fmt)
    sheet.write(curr_row, 2, None, bottom_border_fmt)
    sheet.write(curr_row, 3, None, bottom_border_fmt)
    sheet.write(curr_row, 4, None, lower_right_border_fmt)
    curr_row += 1

    ###
    #
    # CONSULTING Subtable of Breakdown of Charges table
    #
    ###

    # Skip row before Bioinformatics Consulting header.
    sheet.write(curr_row, 1, None, left_border_fmt)
    sheet.write(curr_row, 4, None, right_border_fmt)
    curr_row += 1
    # Write the Bioinformatics Consulting line.
    sheet.write(curr_row, 1, "Bioinformatics Consulting", header_fmt)
    sheet.write(curr_row, 4, None, right_border_fmt)
    curr_row += 1
    # Write the consulting headers.
    sheet.write(curr_row, 1, "Date: Task (Consultant)", col_header_left_fmt)
    sheet.write(curr_row, 2, "Hours (Travel Hours)", col_header_textwrap_fmt)
    sheet.write(curr_row, 3, "Billable Hours", col_header_textwrap_fmt)
    sheet.write(curr_row, 4, "Charge", col_header_right_fmt)
    curr_row += 1

    total_consulting_hours = 0.0
    total_consulting_travel_hours = 0.0
    total_consulting_billable_hours = 0.0
    total_consulting_charges = 0.0

    starting_consulting_row = curr_row

    if len(pi_tag_to_consulting_details[pi_tag]) > 0:

        # Get the rate from the Rates sheet of the BillingConfig workbook.
        rate_consulting_per_hour = get_rates(billing_config_wkbk, 'Bioinformatics Consulting - Stanford')
        rate_consulting_a1_cell  = get_rate_a1_cell(billing_config_wkbk, 'Bioinformatics Consulting - Stanford')

        for (date, summary, consultant, hours, travel_hours, billable_hours) in pi_tag_to_consulting_charges[pi_tag]:

            date_task_consultant_str = "%s: %s (%s)" % (from_excel_date_to_date_string(date), summary, consultant)
            sheet.write(curr_row, 1, date_task_consultant_str, item_entry_textwrap_fmt)

            hours_travel_hours_str = "%s (%s)" % (hours, travel_hours)
            sheet.write(curr_row, 2, hours_travel_hours_str, string_entry_valign_top_fmt)
            sheet.write(curr_row, 3, billable_hours, float_entry_valign_top_fmt)

            charge = rate_consulting_per_hour * billable_hours
            total_consulting_charges += charge

            total_consulting_billable_hours += billable_hours
            total_consulting_hours += hours
            total_consulting_travel_hours += travel_hours

            billable_hours_a1_cell = xl_rowcol_to_cell(curr_row, 3)
            sheet.write_formula(curr_row, 4, '=%s*%s' % (billable_hours_a1_cell, rate_consulting_a1_cell),
                                charge_valign_top_fmt, charge)

            curr_row += 1

    else:
        sheet.write(curr_row, 1, "No consulting", item_entry_fmt)
        sheet.write(curr_row, 4, 0.0, charge_fmt)
        curr_row += 1

    ending_consulting_row = curr_row

    # Skip the line before Total Consulting.
    sheet.write(curr_row, 1, None, left_border_fmt)
    sheet.write(curr_row, 4, None, right_border_fmt)
    curr_row += 1
    # Write the Total Consulting line.
    sheet.write(curr_row, 1, "Total Consulting", bot_header_fmt)
    # sheet.write(curr_row, 2, total_consulting_hours, float_entry_fmt
    # top_hours_a1_cell = xl_rowcol_to_cell(starting_consulting_row, 2)
    # bot_hours_a1_cell = xl_rowcol_to_cell(ending_consulting_row, 2)
    # sheet.write_formula(curr_row, 2, '=SUM(%s:%s)' % (top_hours_a1_cell, bot_hours_a1_cell),
    #                     float_entry_fmt, total_consulting_hours)
    sheet.write(curr_row, 2, "%s (%s)" % (total_consulting_hours, total_consulting_travel_hours), string_entry_fmt)
    # sheet.write(curr_row, 3, total_consulting_billable_hours, float_entry_fmt)
    top_billable_hours_a1_cell = xl_rowcol_to_cell(starting_consulting_row, 3)
    bot_billable_hours_a1_cell = xl_rowcol_to_cell(ending_consulting_row, 3)
    sheet.write_formula(curr_row, 3, '=SUM(%s:%s)' % (top_billable_hours_a1_cell, bot_billable_hours_a1_cell),
                        float_entry_fmt, total_consulting_billable_hours)
    # sheet.write(curr_row, 4, total_consulting_charges, charge_fmt)
    top_charges_a1_cell = xl_rowcol_to_cell(starting_consulting_row, 4)
    bot_charges_a1_cell = xl_rowcol_to_cell(ending_consulting_row, 4)
    sheet.write_formula(curr_row, 4, '=SUM(%s:%s)' % (top_charges_a1_cell, bot_charges_a1_cell),
                        charge_fmt, total_consulting_charges)

    # Save reference to this cell for use in Summary subtable.
    total_consulting_charges_a1_cell = xl_rowcol_to_cell(curr_row, 4)

    curr_row += 1

    # Skip the next line and draw line under this row.
    sheet.write(curr_row, 1, None, lower_left_border_fmt)
    sheet.write(curr_row, 2, None, bottom_border_fmt)
    sheet.write(curr_row, 3, None, bottom_border_fmt)
    sheet.write(curr_row, 4, None, lower_right_border_fmt)
    curr_row += 1

    #
    # Summary of Charges table (B6:E11)
    #

    # Start the Summary of Charges table on the sixth row.
    curr_row = 5
    sheet.write(curr_row, 1, "Summary of Charges:", top_header_fmt)
    sheet.write(curr_row, 2, None, top_border_fmt)
    sheet.write(curr_row, 3, None, top_border_fmt)
    sheet.write(curr_row, 4, None, upper_right_border_fmt)
    curr_row += 1
    # Write the Storage line.
    sheet.write(curr_row, 1, "Storage", header_no_ul_fmt)
    sheet.write(curr_row, 2, total_storage_sizes, float_entry_fmt)
    sheet.write_formula(curr_row, 4, '=%s' % total_storage_charges_a1_cell, big_charge_fmt, total_storage_charges)
    curr_row += 1
    # Write the Computing line.
    sheet.write(curr_row, 1, "Computing", header_no_ul_fmt)
    # sheet.write(curr_row, 2, total_computing_cpuhrs, float_entry_fmt)
    # sheet.write(curr_row, 4, total_computing_charges, big_charge_fmt)
    sheet.write_formula(curr_row, 4, '=%s' % total_computing_charges_a1_cell, big_charge_fmt)
    curr_row += 1
    # Write the Cloud Services line.
    sheet.write(curr_row, 1, "Cloud Services", header_no_ul_fmt)
    sheet.write_formula(curr_row, 4, '=%s' % total_cloud_charges_a1_cell, big_charge_fmt, total_cloud_charges)
    curr_row += 1
    # Write the Consulting line.
    sheet.write(curr_row, 1, "Bioinformatics Consulting", header_no_ul_fmt)
    sheet.write(curr_row, 2, total_consulting_billable_hours, float_entry_fmt)
    sheet.write(curr_row, 4, total_consulting_charges, big_charge_fmt)
    sheet.write_formula(curr_row, 4, '=%s' % total_consulting_charges_a1_cell, big_charge_fmt, total_consulting_charges)
    curr_row += 1
    # Skip a line.
    sheet.write(curr_row, 1, None, left_border_fmt)
    sheet.write(curr_row, 4, None, right_border_fmt)
    curr_row += 1
    # Write the Grand Total line.
    sheet.write(curr_row, 1, "Total Charges", bot_header_border_fmt)
    sheet.write(curr_row, 2, None, bottom_border_fmt)
    sheet.write(curr_row, 3, None, bottom_border_fmt)
    # total_charges = total_storage_charges + total_computing_charges + total_cloud_charges + total_consulting_charges
    sheet.write_formula(curr_row, 4, '=%s+%s+%s+%s' % (total_storage_charges_a1_cell, total_computing_charges_a1_cell, total_cloud_charges_a1_cell, total_consulting_charges_a1_cell),
                        big_bold_charge_fmt)
    curr_row += 1

    #
    # Fill in row in pi_tag -> charges hash.
    #
    # pi_tag_to_charges[pi_tag] = (total_storage_charges, total_computing_charges, total_cloud_charges,
    #                              total_consulting_charges,
    #                              total_charges)


# Copies the Rates sheet from the Rates sheet in the BillingConfig workbook to
# a BillingNotification workbook.
def generate_rates_sheet(rates_input_sheet, rates_output_sheet):

    curr_row = 0
    rates_output_sheet.write(curr_row, 0, "GBSC Rates:", BOLD_FORMAT)
    rates_output_sheet.write(curr_row, 1, "", BOLD_FORMAT)
    rates_output_sheet.write(curr_row, 2, "", BOLD_FORMAT)
    rates_output_sheet.write(curr_row, 3, "", BOLD_FORMAT)

    # Just copy the Rates sheet from the BillingConfig to the BillingNotification.
    curr_row = 1
    for row in range(0, rates_input_sheet.nrows):

        # Read row from input Rates sheet.
        row_values = rates_input_sheet.row_values(row)

        # Write each value from row into output row of output Rates sheet.
        curr_col = 0
        for val in row_values:
            if curr_row == 1:
                rates_output_sheet.write(curr_row, curr_col, val, BOLD_FORMAT)
            elif curr_col == 1:
                rates_output_sheet.write(curr_row, curr_col, val, MONEY_FORMAT)
            else:
                rates_output_sheet.write(curr_row, curr_col, val)
            curr_col += 1
        curr_row += 1


# Generates a Computing Details sheet for a BillingNotification workbook with
# job details associated with a particular PI.  It reads from dict pi_tag_to_job_details.
def generate_computing_details_sheet(sheet, pi_tag):

    # Write the job details, sorted by username.
    curr_row = 1
    for (date, username, job_name, account, node, cpu_core_hrs, jobID, pctage) in sorted(pi_tag_to_job_details[pi_tag],key=lambda s: s[1]):

        curr_col = 0
        sheet.write(curr_row, curr_col, date, DATE_FORMAT) ; curr_col += 1
        sheet.write(curr_row, curr_col, username) ; curr_col += 1
        sheet.write(curr_row, curr_col, job_name) ; curr_col += 1
        sheet.write(curr_row, curr_col, account) ; curr_col += 1
        sheet.write(curr_row, curr_col, node) ; curr_col += 1
        sheet.write(curr_row, curr_col, cpu_core_hrs, FLOAT_FORMAT) ; curr_col += 1
        sheet.write(curr_row, curr_col, jobID) ; curr_col += 1
        sheet.write(curr_row, curr_col, pctage, PERCENT_FORMAT) ; curr_col += 1

        # Advance to the next row.
        curr_row += 1


# Generates the Lab Users sheet for a BillingNotification workbook with
# username details for a particular PI.  It reads from dicts:
#  cloud_project_account_to_cloud_details
#  pi_tag_to_cloud_account_pctages
def generate_cloud_details_sheet(sheet, pi_tag):

    curr_row = 1

    # Get the list of accounts associated with this PI.
    for (account, pctage) in pi_tag_to_cloud_account_pctages[pi_tag]:

        for project in cloud_account_to_cloud_projects[account]:

            # Write the cloud details.
            for (platform, description, dates, quantity, uom, charge) in cloud_project_account_to_cloud_details[(project, account)]:

                curr_col = 0
                sheet.write(curr_row, curr_col, platform) ; curr_col += 1
                # If we have the project number here, use the project name.
                if len(project) > 0 and project[0].isdigit():
                    sheet.write(curr_row, curr_col, cloud_projnum_to_cloud_project[project]) ; curr_col += 1
                else:
                    sheet.write(curr_row, curr_col, project) ; curr_col += 1
                sheet.write(curr_row, curr_col, description) ; curr_col += 1
                sheet.write(curr_row, curr_col, dates) ; curr_col += 1
                sheet.write(curr_row, curr_col, quantity, FLOAT_FORMAT) ; curr_col += 1
                sheet.write(curr_row, curr_col, uom) ; curr_col += 1
                sheet.write(curr_row, curr_col, charge, MONEY_FORMAT) ; curr_col += 1
                sheet.write(curr_row, curr_col, pctage, PERCENT_FORMAT) ; curr_col += 1

                lab_cost = charge * pctage
                sheet.write(curr_row, curr_col, lab_cost, MONEY_FORMAT) ; curr_col += 1

                # Advance to the next row.
                curr_row += 1


# Generates the Consulting Details sheet for a BillingNotifications workbook with
# consulting details for a particular PI.  It reads from dict:
#  pi_tag_to_consulting_details
def generate_consulting_details_sheet(sheet, pi_tag):

    curr_row = 1  # The header is already in this sheet

    for (date, summary, notes, consultant, hours, travel_hours, cumul_hours) in pi_tag_to_consulting_details[pi_tag]:

        curr_col = 0
        sheet.write(curr_row, curr_col, date, DATE_FORMAT) ; curr_col += 1
        sheet.write(curr_row, curr_col, summary) ; curr_col += 1
        sheet.write(curr_row, curr_col, notes) ; curr_col += 1
        sheet.write(curr_row, curr_col, consultant) ; curr_col += 1
        sheet.write(curr_row, curr_col, hours, FLOAT_FORMAT) ; curr_col += 1
        sheet.write(curr_row, curr_col, travel_hours, FLOAT_FORMAT) ; curr_col += 1
        sheet.write(curr_row, curr_col, cumul_hours, FLOAT_FORMAT) ; curr_col += 1

        curr_row += 1



# Generates the Lab Users sheet for a BillingNotification workbook with
# username details for a particular PI.  It reads from dicts pi_tag_to_user_details and username_to_user_details.
def generate_lab_users_sheet(sheet, pi_tag):

    # Write the user details for active users and moving the inactive users to a separate list.
    past_user_details = []
    curr_row = 1
    for (username, date_added, date_removed, pctage) in pi_tag_to_user_details[pi_tag]:

        curr_col = 0

        # Get the user details for username.
        (email, fullname) = username_to_user_details[username]

        if date_removed == '':
            sheet.write(curr_row, curr_col, username) ; curr_col += 1
            sheet.write(curr_row, curr_col, fullname) ; curr_col += 1
            sheet.write(curr_row, curr_col, email) ; curr_col += 1
            sheet.write(curr_row, curr_col, date_added, DATE_FORMAT) ; curr_col += 1
            sheet.write(curr_row, curr_col, "current") ; curr_col += 1
            curr_row += 1
        else:
            # Users who have been removed will be listed in a table below the current lab members
            past_user_details.append([username, email, fullname, date_added, date_removed])

    # Write out a subheader for the Previous Lab Members.
    curr_row += 1  # Skip a row before the subheader.
    sheet.write(curr_row, 0, "Previous Lab Members", BOLD_FORMAT)
    curr_row += 1
    for (username, email, fullname, date_added, date_removed) in past_user_details:

        curr_col = 0
        sheet.write(curr_row, curr_col, username) ; curr_col += 1
        sheet.write(curr_row, curr_col, fullname) ; curr_col += 1
        sheet.write(curr_row, curr_col, email) ; curr_col += 1
        sheet.write(curr_row, curr_col, date_added, DATE_FORMAT) ; curr_col += 1
        sheet.write(curr_row, curr_col, date_removed, DATE_FORMAT) ; curr_col += 1

        curr_row += 1


# Generates the Totals sheet for a BillingAggregate workbook, populating the sheet
# from the pi_tag_to_charges dict.
def generate_aggregrate_sheet(sheet):

    # Set column widths
    sheet.set_column("A:A", 12)
    sheet.set_column("B:B", 12)
    sheet.set_column("C:C", 12)
    sheet.set_column("D:D", 12)
    sheet.set_column("E:E", 12)
    sheet.set_column("F:F", 12)
    sheet.set_column("G:G", 12)
    sheet.set_column("H:H", 12)

    total_fmt = make_format(billing_aggreg_wkbk,
                            {'font_size': 14, 'bold': True})
    charge_fmt = make_format(billing_aggreg_wkbk,
                             {'font_size': 10, 'align': 'right',
                              'num_format': '$#,##0.00'})
    sub_total_charge_fmt = make_format(billing_aggreg_wkbk,
                                       {'font_size': 14, 'align': 'right',
                                        'num_format': '$#,##0.00'})
    grand_charge_fmt = make_format(billing_aggreg_wkbk,
                                   {'font_size': 14, 'align': 'right', 'bold': True,
                                    'num_format': '$#,##0.00'})

    sub_total_storage = 0.0
    sub_total_computing = 0.0
    sub_total_cloud = 0.0
    sub_total_consulting = 0.0
    grand_total_charges = 0.0

    # Compute column numbers for various columns.
    storage_column_num     = BILLING_AGGREG_SHEET_COLUMNS['Totals'].index('Storage')
    computing_column_num   = BILLING_AGGREG_SHEET_COLUMNS['Totals'].index('Computing')
    cloud_column_num       = BILLING_AGGREG_SHEET_COLUMNS['Totals'].index('Cloud')
    consulting_column_num  = BILLING_AGGREG_SHEET_COLUMNS['Totals'].index('Consulting')

    curr_row = 1
    for pi_tag in sorted(pi_tag_to_charges.iterkeys()):

        (storage, computing, cloud, consulting, total_charges) = pi_tag_to_charges[pi_tag]
        (pi_first_name, pi_last_name, _) = pi_tag_to_names_email[pi_tag]

        curr_col = 0
        sheet.write(curr_row, curr_col, pi_first_name) ; curr_col += 1
        sheet.write(curr_row, curr_col, pi_last_name) ; curr_col += 1
        sheet.write(curr_row, curr_col, pi_tag) ; curr_col += 1
        sheet.write(curr_row, curr_col, storage, charge_fmt) ; curr_col += 1
        sheet.write(curr_row, curr_col, computing, charge_fmt) ; curr_col += 1
        sheet.write(curr_row, curr_col, cloud, charge_fmt) ; curr_col += 1
        sheet.write(curr_row, curr_col, consulting, charge_fmt) ; curr_col += 1

        storage_a1_cell = xl_rowcol_to_cell(curr_row, storage_column_num)
        computing_a1_cell = xl_rowcol_to_cell(curr_row, computing_column_num)
        cloud_a1_cell = xl_rowcol_to_cell(curr_row, cloud_column_num)
        consulting_a1_cell = xl_rowcol_to_cell(curr_row, consulting_column_num)


        sheet.write_formula(curr_row, curr_col, '=SUM(%s:%s)' % (storage_a1_cell, consulting_a1_cell),
                            charge_fmt, total_charges) ; curr_col += 1

        sub_total_storage += storage
        sub_total_computing += computing
        sub_total_cloud += cloud
        sub_total_consulting += consulting
        grand_total_charges += total_charges

        curr_row += 1

    storage_a1_cell = xl_rowcol_to_cell(curr_row, storage_column_num)
    computing_a1_cell = xl_rowcol_to_cell(curr_row, computing_column_num)
    cloud_a1_cell = xl_rowcol_to_cell(curr_row, cloud_column_num)
    consulting_a1_cell = xl_rowcol_to_cell(curr_row, consulting_column_num)

    sheet.write(curr_row, 0, "TOTALS", total_fmt)
    #sheet.write(curr_row, storage_column_num, sub_total_storage, sub_total_charge_fmt)
    top_storage_a1_cell = xl_rowcol_to_cell(1, storage_column_num)
    bot_storage_a1_cell = xl_rowcol_to_cell(curr_row - 1, storage_column_num)
    sheet.write_formula(curr_row, storage_column_num, '=SUM(%s:%s)' % (top_storage_a1_cell, bot_storage_a1_cell),
                        sub_total_charge_fmt, sub_total_storage)
    #sheet.write(curr_row, computing_column_num, sub_total_computing, sub_total_charge_fmt)
    top_computing_a1_cell = xl_rowcol_to_cell(1, computing_column_num)
    bot_computing_a1_cell = xl_rowcol_to_cell(curr_row - 1, computing_column_num)
    sheet.write_formula(curr_row, computing_column_num, '=SUM(%s:%s)' % (top_computing_a1_cell, bot_computing_a1_cell),
                        sub_total_charge_fmt, sub_total_computing)

    top_cloud_a1_cell = xl_rowcol_to_cell(1, cloud_column_num)
    bot_cloud_a1_cell = xl_rowcol_to_cell(curr_row - 1, cloud_column_num)
    sheet.write_formula(curr_row, cloud_column_num, '=SUM(%s:%s)' % (top_cloud_a1_cell, bot_cloud_a1_cell),
                        sub_total_charge_fmt, sub_total_cloud)

    #sheet.write(curr_row, consulting_column_num, sub_total_consulting, sub_total_charge_fmt)
    top_consulting_a1_cell = xl_rowcol_to_cell(1, consulting_column_num)
    bot_consulting_a1_cell = xl_rowcol_to_cell(curr_row - 1, consulting_column_num)
    sheet.write_formula(curr_row, consulting_column_num, '=SUM(%s:%s)' % (top_consulting_a1_cell, bot_consulting_a1_cell),
                        sub_total_charge_fmt, sub_total_consulting)

    sheet.write_formula(curr_row, consulting_column_num+1, '=%s+%s+%s+%s' % (storage_a1_cell, computing_a1_cell, cloud_a1_cell, consulting_a1_cell),
                        grand_charge_fmt, grand_total_charges)

#=====
#
# SCRIPT BODY
#
#=====

parser = argparse.ArgumentParser()

parser.add_argument("billing_config_file",
                    help='The BillingConfig file')
parser.add_argument("-d","--billing_details_file",
                    default=None,
                    help='The BillingDetails file')
parser.add_argument("-r", "--billing_root",
                    default=None,
                    help='The Billing Root directory [default = None]')
parser.add_argument("-p", "--pi_sheets", action="store_true",
                    default=False,
                    help='Add PI-specific sheets to the BillingAggregate workbook [default = False]')
parser.add_argument("-v", "--verbose", action="store_true",
                    default=False,
                    help='Get real chatty [default = false]')
parser.add_argument("-y","--year", type=int, choices=range(2013,2031),
                    default=None,
                    help="The year to be used. [default = this year]")
parser.add_argument("-m", "--month", type=int, choices=range(1,13),
                    default=None,
                    help="The month to be used. [default = last month]")

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

# Get the absolute path for the billing_config_file.
billing_config_file = os.path.abspath(args.billing_config_file)

###
#
# Read the BillingConfig workbook and build input data structures.
#
###

billing_config_wkbk = xlrd.open_workbook(billing_config_file)

#
# Get the location of the BillingRoot directory from the Config sheet.
#
(billing_root, _) = read_config_sheet(billing_config_wkbk)

# Override billing_root with switch args, if present.
if args.billing_root is not None:
    billing_root = args.billing_root
# If we still don't have a billing root dir, use the current directory.
if billing_root is None:
    billing_root = os.getcwd()

# Get the absolute path for the billing_root directory.
billing_root = os.path.abspath(billing_root)

# Within BillingRoot, create YEAR/MONTH dirs if necessary.
year_month_dir = os.path.join(billing_root, str(year), "%02d" % month)
if not os.path.exists(year_month_dir):
    os.makedirs(year_month_dir)

# If BillingDetails file given, use that, else look in BillingRoot.
if args.billing_details_file is not None:
    billing_details_file = args.billing_details_file
else:
    billing_details_file = os.path.join(year_month_dir, "%s.%s-%02d.xlsx" % (BILLING_DETAILS_PREFIX, year, month))

# Get the absolute path for the billing_details_file.
billing_details_file = os.path.abspath(billing_details_file)

#
# Output the state of arguments.
#
print "GENERATING NOTIFICATIONS FOR %02d/%d:" % (month, year)
print "  BillingConfigFile: %s" % (billing_config_file)
print "  BillingRoot: %s" % (billing_root)
print "  BillingDetailsFile: %s" % (billing_details_file)
print

#
# Build data structures.
#
print "Building configuration data structures."
build_global_data(billing_config_wkbk, begin_month_timestamp, end_month_timestamp)

###
#
# Read the BillingDetails workbook, and create output data structures.
#
###

# Open the BillingDetails workbook.
print "Read in BillingDetails workbook."
billing_details_wkbk = xlrd.open_workbook(billing_details_file)

# Read in its Storage sheet and generate output data.
print "Reading storage sheet."
read_storage_sheet(billing_details_wkbk)

# Read in its Computing sheet and generate output data.
print "Reading computing sheet."
read_computing_sheet(billing_details_wkbk)

print "Reading cloud sheet."
read_cloud_sheet(billing_details_wkbk)

# Read in its Consulting sheet and generate output data.
print "Reading consulting sheet."
read_consulting_sheet(billing_details_wkbk)

###
#
# Write BillingNotification workbooks from output data structures.
#
###

print "Writing notification workbooks:"
for pi_tag in sorted(pi_tag_list):

    print " %s" % pi_tag

    # Initialize the BillingNotification spreadsheet for this PI.
    notifs_wkbk_filename = "%s-%s.%s-%02d.xlsx" % (BILLING_NOTIFS_PREFIX, pi_tag, year, month)
    notifs_wkbk_pathname = os.path.join(year_month_dir, notifs_wkbk_filename)

    billing_notifs_wkbk = xlsxwriter.Workbook(notifs_wkbk_pathname)
    sheet_name_to_sheet_map = init_billing_notifs_wkbk(billing_notifs_wkbk)

    # Generate the Billing sheet.
    generate_billing_sheet(billing_notifs_wkbk, sheet_name_to_sheet_map['Billing'],
                           pi_tag, begin_month_timestamp, end_month_timestamp)

    # Generate the Rates sheet.
    generate_rates_sheet(billing_config_wkbk.sheet_by_name('Rates'), sheet_name_to_sheet_map['Rates'])

    # Generate the Computing Details sheet.
    generate_computing_details_sheet(sheet_name_to_sheet_map['Computing Details'], pi_tag)

    # Generate the Cloud Details sheet.
    generate_cloud_details_sheet(sheet_name_to_sheet_map['Cloud Details'], pi_tag)

    # Generate the Lab Users sheet.
    generate_lab_users_sheet(sheet_name_to_sheet_map['Lab Users'], pi_tag)

    # Generate the Consulting Details
    generate_consulting_details_sheet(sheet_name_to_sheet_map['Consulting Details'], pi_tag)

    billing_notifs_wkbk.close()

###
#
# Write BillingAggregate workbook from totals in BillingNotifications workbooks.
#
###

print "Writing billing aggregate workbook."

aggreg_wkbk_filename = "%s.%s-%02d.xlsx" % (BILLING_NOTIFS_PREFIX, year, month)
aggreg_wkbk_pathname = os.path.join(year_month_dir, aggreg_wkbk_filename)

billing_aggreg_wkbk = xlsxwriter.Workbook(aggreg_wkbk_pathname)

aggreg_sheet_name_to_sheet = init_billing_aggreg_wkbk(billing_aggreg_wkbk, pi_tag_list)

aggreg_totals_sheet = aggreg_sheet_name_to_sheet['Totals']

# Create the aggregate Totals sheet.
generate_aggregrate_sheet(aggreg_totals_sheet)

if args.pi_sheets:
    # Add the Billing sheets for each PI.
    for pi_tag in sorted(pi_tag_list):

        pi_sheet = aggreg_sheet_name_to_sheet[pi_tag]

        generate_billing_sheet(billing_aggreg_wkbk, pi_sheet,
                               pi_tag, begin_month_timestamp, end_month_timestamp)

billing_aggreg_wkbk.close()

###
#
# Output some summary statistics.
#
###
total_jobs_billed = 0
for pi_tag in pi_tag_list:
    total_jobs_billed += len(pi_tag_to_job_details[pi_tag])

print "Total Jobs Billed:", total_jobs_billed