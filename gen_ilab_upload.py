#!/usr/bin/env python

#===============================================================================
#
# gen_ilab_upload.py - Generate billing data for upload into iLab.
#
# ARGS:
#   1st: the BillingConfig spreadsheet.
#
# SWITCHES:
#   --billing_details_file: Location of the BillingDetails.xlsx file (default=look in BillingRoot/<year>/<month>)
#   --billing_root:    Location of BillingRoot directory (overrides BillingConfig.xlsx)
#                      [default if no BillingRoot in BillingConfig.xlsx or switch given: CWD]
#   --year:            Year of snapshot requested. [Default is this year]
#   --month:           Month of snapshot requested. [Default is last month]
#
# OUTPUT:
#   CSV file with billing data suitable for uploading into iLab.
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
import codecs
from collections import defaultdict
import csv
import datetime
import locale  # for converting strings with commas into floats
import os
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
# From billing_common.py
global BILLING_NOTIFS_SHEET_COLUMNS
global BILLING_AGGREG_SHEET_COLUMNS
global BILLING_DETAILS_PREFIX
global BILLING_NOTIFS_PREFIX
global GOOGLE_INVOICE_PREFIX
global ILAB_EXPORT_PREFIX
global CONSULTING_HOURS_FREE
global CONSULTING_TRAVEL_RATE_DISCOUNT

# Default headers for the ilab Export CSV file (if not read in from iLab template file).
DEFAULT_CSV_HEADERS = ['service_id','note','service_quantity','purchased_on',
                       'service_request_id','owner_email','pi_email']

# Default available services table (to be used if no available services file given).
DEFAULT_AVAILABLE_SERVICES_ID_DICT = {
    'Local Storage'   : ['Local Cluster Storage', 1991],
    'Local Computing' : ['Local Cluster Computing', 1992],
    'Cloud Services' : ['Cloud Services (Passthrough)', 2355],
    'Consulting Free' : ['Consulting - First 1 hour (Units are hours)', 2349],
    'Consulting Paid' : ['Consulting - Work on a Project beyond 3 hours (Units are hours)', 2350],
    'Consulting Resources' : ['Consulting Compute Cost (Passthrough)', 2356]
}

# Available Service ID names (for reading in available service info from iLab Available Services file).
AVAILABLE_SERVICES_COLUMN_NAME       = 'Name'
AVAILABLE_SERVICES_COLUMN_SERVICE_ID = 'Service ID'

#=====
#
# GLOBALS
#
#=====

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

# Mapping from pi_tags to iLab service request IDs (1-to-1 mapping).
pi_tag_to_ilab_service_req_id = dict()

# Mapping from job_tags to list of [pi_tag, %age].
job_tag_to_pi_tag_pctages = defaultdict(list)

# Mapping from folders to list of [pi_tag, %age].
folder_to_pi_tag_pctages = defaultdict(list)

# Mapping from pi_tag to list of [folder, size, %age].
pi_tag_to_folder_sizes = defaultdict(list)

# Mapping from pi_tag to list of [username, cpu_core_hrs, %age].
pi_tag_to_username_cpus = defaultdict(list)

# Mapping from pi_tag to list of [job_tag, cpu_core_hrs, %age].
pi_tag_to_job_tag_cpus = defaultdict(list)

# Mapping from pi_tag to list of [date, username, job_name, account, cpu_core_hrs, jobID, %age].
pi_tag_to_sge_job_details = defaultdict(list)

# Mapping from pi_tag to list of [username, date_added, date_removed, %age].
pi_tag_to_user_details = defaultdict(list)

# Mapping from pi_tag to set of (cloud account, %age) tuples.
pi_tag_to_cloud_account_pctages = defaultdict(set)

# Mapping from cloud account to set of cloud projects.
cloud_account_to_cloud_projects = defaultdict(set)

# Mapping from cloud project to lists of (platform, account, description, dates, quantity, UOM, charge) tuples.
cloud_project_account_to_cloud_details = defaultdict(list)

# Mapping from cloud project to total charge.
cloud_project_account_to_total_charges = defaultdict(float)

# Mapping from pi_tag to list of (date, summary, hours, cumul_hours)
consulting_details = defaultdict(list)


# Set locale to be US english for converting strings with commas into floats.
locale.setlocale(locale.LC_ALL, 'en_US.UTF-8')

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
global from_date_string_to_timestamp
global sheet_get_named_column
global read_config_sheet
global config_sheet_get_dict

# Filters a list of lists using a parallel list of [date_added, date_removed]'s.
# Returns the elements in the first list which are valid with the month date range given.
def filter_by_dates(obj_list, date_list, begin_month_exceldate, end_month_exceldate):

    output_list = []

    for (obj, (date_added, date_removed)) in zip(obj_list, date_list):

        # If the date added is BEFORE the end of this month, and
        #    the date removed is AFTER the beginning of this month,
        # then save the account information in the mappings.
        if date_added < end_month_exceldate and date_removed >= begin_month_exceldate:
            output_list.append(obj)

    return output_list


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

#
# Reads a subtable from the CSVFile file-object, which is all the lines
# between blank lines.
#
def get_google_invoice_csv_subtable_lines(csvfile_obj):

    subtable = []

    line = csvfile_obj.readline()
    while not line.startswith(',') and line != '' and line != '\n':
        subtable.append(line)
        line = csvfile_obj.readline()

    return subtable


# Creates all the data structures used to write the BillingNotification workbook.
# The overall goal is to mimic the tables of the notification sheets so that
# to build the table, all that is needed is to print out one of these data structures.
def build_global_data(wkbk, begin_month_timestamp, end_month_timestamp, read_cloud_data):

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
    # Create mapping from pi_tag to iLab Service Request ID.
    #
    global pi_tag_to_ilab_service_req_id

    pi_ilab_ids = sheet_get_named_column(pis_sheet,"iLab Service Request ID")

    pi_tag_to_ilab_service_req_id = dict(zip(pi_tag_list, pi_ilab_ids))

    # Organize data from the Cloud sheet, if present.
    if read_cloud_data:
        cloud_sheet = wkbk.sheet_by_name("Cloud")

        #
        # Create mapping from pi_tag to cloud project from the BillingConfig PIs sheet.
        # Create mapping from cloud project to list of (pi_tag, %age) tuples.
        # Create mapping from cloud project to cloud account (1-to-1).
        #
        global pi_tag_to_cloud_account_pctages
        global cloud_account_to_cloud_projects

        cloud_pi_tags     = sheet_get_named_column(cloud_sheet, "PI Tag")
        cloud_projects    = sheet_get_named_column(cloud_sheet, "Project")
        cloud_projnums    = sheet_get_named_column(cloud_sheet, "Project Number")
        cloud_accounts    = sheet_get_named_column(cloud_sheet, "Account")
        cloud_pctage      = sheet_get_named_column(cloud_sheet, "%age")

        cloud_dates_added = sheet_get_named_column(cloud_sheet, "Date Added")
        cloud_dates_remvd = sheet_get_named_column(cloud_sheet, "Date Removed")

        cloud_rows = filter_by_dates(zip(cloud_pi_tags, cloud_projects, cloud_projnums,
                                     cloud_accounts, cloud_pctage),
                                     zip(cloud_dates_added, cloud_dates_remvd),
                                     begin_month_exceldate, end_month_exceldate)

        for (pi_tag, project, projnum, account, pctage) in cloud_rows:

            # Associate the project name and percentage to be charged with the pi_tag.
            pi_tag_to_cloud_account_pctages[pi_tag].add((account, pctage))

            # Associate the project number with the pi_tag also, in case the project is deleted and loses its name.
            #pi_tag_to_cloud_account_pctages[pi_tag].append((projnum, pctage))

            # Associate the account with the project name and with the project number.
            cloud_account_to_cloud_projects[account].add(project)
            cloud_account_to_cloud_projects[account].add(projnum)
            cloud_account_to_cloud_projects[account].add("")  # For credits to the account.

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
        folder_to_pi_tag_pctages[folder].append([pi_tag, pctage])


# Reads the Storage sheet of the BillingDetails workbook given, and populates
# the pi_tag_to_folder_sizes dict with the folder measurements for each PI.
def read_storage_sheet(storage_sheet):

    global pi_tag_to_folder_sizes

    for row in range(1,storage_sheet.nrows):

        (date, timestamp, folder, size, used) = storage_sheet.row_values(row)

        # List of [pi_tag, %age] pairs.
        pi_tag_pctages = folder_to_pi_tag_pctages[folder]

        for (pi_tag, pctage) in pi_tag_pctages:
            pi_tag_to_folder_sizes[pi_tag].append([folder, size, pctage])


# Reads the Computing sheet of the BillingDetails workbook given, and populates
# the job_tag_to_pi_tag_cpus, pi_tag_to_job_tag_cpus, pi_tag_to_username_cpus, and
# pi_tag_to_sge_job_details dicts.
def read_computing_sheet(computing_sheet):

    global pi_tag_to_sge_job_details
    global pi_tag_to_job_tag_cpus
    global pi_tag_to_username_cpus

    for row in range(1,computing_sheet.nrows):

        (job_date, job_timestamp, job_username, job_name, account, node, cores, wallclock, jobID) = \
            computing_sheet.row_values(row)

        # Calculate CPU-core-hrs for job.
        cpu_core_hrs = cores * wallclock / 3600.0  # wallclock is in seconds.

        # Rename this variable for easier understanding.
        job_tag = account.lower()

        # If there is a job_tag in the account field and the job tag is known, credit the job_tag with the job CPU time.
        # Else, credit the user with the job.
        if (job_tag != '' and
            (job_tag_to_pi_tag_pctages.get(job_tag) is not None or job_tag.lower() in pi_tag_list)):

            # All PIs have a default job_tag that can be applied to jobs to be billed to them.
            if job_tag.lower() in pi_tag_list:
                job_tag = job_tag.lower()
                job_pi_tag_pctage_list = [[job_tag, 1.0]]
            else:
                job_pi_tag_pctage_list = job_tag_to_pi_tag_pctages[job_tag]

            # If no pi_tag is associated with this job tag, speak up.
            if len(job_pi_tag_pctage_list) == 0:
                print "   No PI associated with job ID %s" % jobID

            # Distribute this job's CPU-hrs amongst pi_tags by %ages.
            for (pi_tag, pctage) in job_pi_tag_pctage_list:

                 # This list is (job_tag, cpu_core_hrs, %age).
                 job_tag_cpu_list = pi_tag_to_job_tag_cpus.get(pi_tag)

                 # If pi_tag has an existing list of job_tag/CPUs:
                 if job_tag_cpu_list is not None:

                     # Find if job_tag is in list of job_tag/CPUs for this pi_tag.
                     for job_tag_cpu in job_tag_cpu_list:
                         pi_job_tag = job_tag_cpu[0]

                         # Increment the job_tag's CPUs.
                         if pi_job_tag == job_tag:
                             job_tag_cpu[1] += cpu_core_hrs
                             break
                     else:
                         # No matching job_tag in pi_tag list -- add a new one to the list.
                         job_tag_cpu_list.append([job_tag, cpu_core_hrs, pctage])

                 # Else start a new job_tag/CPUs list for the pi_tag.
                 else:
                     pi_tag_to_job_tag_cpus[pi_tag] = [[job_tag, cpu_core_hrs, pctage]]

                 #
                 # Save job details for pi_tag.
                 #
                 new_job_details = [job_date, job_username, job_name, account, cpu_core_hrs, jobID, pctage]
                 pi_tag_to_sge_job_details[pi_tag].append(new_job_details)

        # Else credit a user with the job CPU time.
        else:
            pi_tag_pctages = get_pi_tags_for_username_by_date(job_username, job_timestamp)

            if len(pi_tag_pctages) == 0:
                print "   No PI associated with user %s" % job_username

            for (pi_tag, pctage) in pi_tag_pctages:

                # if pctage == 0.0: continue

                #
                # Increment this user's CPU-core-hrs.
                #

                # This list is (username, cpu_core_hrs, %age).
                username_cpu_list = pi_tag_to_username_cpus.get(pi_tag)

                # If pi_tag has an existing list of user/CPUs:
                if username_cpu_list is not None:

                    # Find if job_username is in list of user/CPUs for this pi_tag.
                    for username_cpu in username_cpu_list:
                        username = username_cpu[0]
                        user_pctage = username_cpu[2]

                        # Increment the user's CPUs if they already exist in the list.
                        if username == job_username and user_pctage == pctage:
                            username_cpu[1] += cpu_core_hrs
                            break
                    else:
                        # No matching user in pi_tag list -- add a new one to the list.
                        username_cpu_list.append([job_username, cpu_core_hrs, pctage])

                # Else start a new user/CPUs list for the pi_tag.
                else:
                    pi_tag_to_username_cpus[pi_tag] = [[job_username, cpu_core_hrs, pctage]]

                #
                # Save job details for pi_tag.
                #
                new_job_details = [job_date, job_username, job_name, account, cpu_core_hrs, jobID, pctage]
                pi_tag_to_sge_job_details[pi_tag].append(new_job_details)


# Read the Cloud sheet from the BillingDetails workbook.
def read_cloud_sheet(cloud_sheet):

    for row in range(1,cloud_sheet.nrows):

        (platform, account, project, description, dates, quantity, uom, charge) = cloud_sheet.row_values(row)

        # If project is of the form "<project-name> (<project-id>)", remove the "(<project-id>)"
        project_id_index = project.find(" (")
        if project_id_index != -1:
            project = project[:project_id_index]

        # Save the cloud item in a list of charges for that PI.
        cloud_project_account_to_cloud_details[(project, account)].append((platform, description, dates, quantity, uom, charge))

        # Accumulate the total cost of a project.
        cloud_project_account_to_total_charges[(project, account)] += float(charge)


def read_google_invoice(google_invoice_csv_file):

    ###
    # Read the Google Invoice CSV File
    ###

    # Google Invoice CSV files are Unicode with BOM.
    google_invoice_csv_file_obj = codecs.open(google_invoice_csv_file, 'rU', encoding='utf-8-sig')

    #  Read the header subtable
    google_invoice_header_subtable = get_google_invoice_csv_subtable_lines(google_invoice_csv_file_obj)

    google_invoice_header_csvreader = csv.DictReader(google_invoice_header_subtable, fieldnames=['key', 'value'])

    for row in google_invoice_header_csvreader:

        #   Extract invoice date from "Issue Date".
        if row['key'] == 'Issue date':
            google_invoice_issue_date = row['value']
        #   Extract the "Amount Due" value.
        elif row['key'] == 'Amount due':
            google_invoice_amount_due = locale.atof(row['value'])

    print >> sys.stderr, "  Amount due: $%0.2f" % (google_invoice_amount_due)

    # Accumulate the total amount of charges while processing each line,
    #  to compare with total amount in header.
    google_invoice_total_amount = 0.0

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

            #     Accumulate total charges.
            amount = locale.atof(row_dict['Amount'])
            google_invoice_total_amount += amount

            google_account = row_dict['Order']

            #     Construct note for ilab entry.
            google_platform = 'Google Cloud Platform, Firebase, and APIs'
            google_project = row_dict['Source']
            google_item    = row_dict['Description']
            google_quantity = row_dict['Quantity']
            google_uom     = row_dict['UOM']
            google_dates   = row_dict['Interval']

            # Save the cloud details with the appropriate PI.
            cloud_project_account_to_cloud_details[(google_project, google_account)].append((google_platform, google_item, google_dates,
                                                                                             google_quantity, google_uom, amount))

    # Compare total charges to "Amount Due".
    if abs(google_invoice_total_amount - google_invoice_amount_due) >= 0.01:  # Ignore differences less than a penny.
        print >> sys.stderr, "  WARNING: Accumulated amounts do not equal amount due: ($%.2f != $%.2f)" % (google_invoice_total_amount,
                                                                                                           google_invoice_amount_due)
    else:
        print >> sys.stderr, "  VERIFIED: Sum of individual transactions equals Amount due."


#
# Read in the Consulting sheet.
#
# It fills in the dict consulting_details.
#
def read_consulting_sheet(consulting_sheet):

    for row in range(1, consulting_sheet.nrows):

        (date, pi_tag, hours, travel_hours, participants, summary, notes, cumul_hours) = consulting_sheet.row_values(row)

        # Save the consulting item in a list of charges for that PI.
        consulting_details[pi_tag].append((date, summary, float(hours), float(travel_hours), float(cumul_hours)))

#
# Digest cluster data and output Cluster iLab file.
#
def process_cluster_data():

    # Read in its Storage sheet.
    print "Reading storage sheet."
    storage_sheet = billing_details_wkbk.sheet_by_name("Storage")
    read_storage_sheet(storage_sheet)

    # Read in its Computing sheet.
    print "Reading computing sheet."
    computing_sheet = billing_details_wkbk.sheet_by_name("Computing")
    read_computing_sheet(computing_sheet)

    ###
    #
    # Write iLab export CSV file from output data structures.
    #
    ###
    print "Writing out BillingDetails lines for Cluster into iLab export CSV file."

    ###
    #
    # Open the iLab CSV file for writing out.
    #
    ###

    ilab_export_csv_filename = "%s-Cluster.%s-%02d.csv" % (ILAB_EXPORT_PREFIX, year, month)
    ilab_export_csv_pathname = os.path.join(year_month_dir, ilab_export_csv_filename)

    ilab_export_csv_file = open(ilab_export_csv_pathname, "w")

    ilab_export_csv_dictwriter = csv.DictWriter(ilab_export_csv_file, fieldnames=ilab_csv_headers)

    ilab_export_csv_dictwriter.writeheader()

    # Write out cluster data to iLab export CSV file.
    for pi_tag in sorted(pi_tag_list):
        print " %s" % pi_tag

        _ = output_ilab_csv_data_for_cluster(ilab_export_csv_dictwriter, pi_tag,
                                             ilab_service_id_local_storage, ilab_service_id_local_computing,
                                             begin_month_timestamp, end_month_timestamp)

    # Close the iLab export CSV file.
    ilab_export_csv_file.close()

#
# Digest cloud data and output Cloud iLab file.
#
def process_cloud_data():

    # Read in Cloud data from Google Invoice, if given as argument.
    if args.google_invoice_csv is not None:

        ###
        # Read in Google Cloud Invoice data, ignoring data from BillingDetails.
        ###
        print "Reading Google Invoice."
        read_google_invoice(google_invoice_csv)

    # Read in the Cloud sheet from the BillingDetails file, if present.
    elif "Cloud" in billing_details_wkbk.sheet_names():

        print "Reading cloud sheet."
        cloud_sheet = billing_details_wkbk.sheet_by_name("Cloud")
        read_cloud_sheet(cloud_sheet)

    else:
        print "No Cloud sheet in BillingDetails nor Google Invoice file...skipping"
        return

    print "Writing out Cloud details into iLab export CSV file."

    # Open the iLab CSV file for writing out.
    ilab_export_csv_filename = "%s-Cloud.%s-%02d.csv" % (ILAB_EXPORT_PREFIX, year, month)
    ilab_export_csv_pathname = os.path.join(year_month_dir, ilab_export_csv_filename)

    ilab_export_csv_file = open(ilab_export_csv_pathname, "w")

    ilab_export_csv_dictwriter = csv.DictWriter(ilab_export_csv_file, fieldnames=ilab_csv_headers)

    ilab_export_csv_dictwriter.writeheader()

    for pi_tag in pi_tag_list:
        print " %s" % pi_tag

        ret_val = output_ilab_csv_data_for_cloud(ilab_export_csv_dictwriter, pi_tag, ilab_service_id_google_passthrough,
                                                 begin_month_timestamp, end_month_timestamp)

    # Close the iLab export CSV file.
    ilab_export_csv_file.close()

#
# Digest Consulting data and output Consulting iLab file.
#
def process_consulting_data():

    # Read in its Consulting sheet.
    if "Consulting" in billing_details_wkbk.sheet_names():
        print "Reading consulting sheet."
        consulting_sheet = billing_details_wkbk.sheet_by_name("Consulting")
        read_consulting_sheet(consulting_sheet)
    else:
        print "No consulting sheet in BillingDetails: skipping"
        return

    print "Writing out Consulting details into iLab export CSV file."

    # Open the iLab CSV file for writing out.
    ilab_export_csv_filename = "%s-Consulting.%s-%02d.csv" % (ILAB_EXPORT_PREFIX, year, month)
    ilab_export_csv_pathname = os.path.join(year_month_dir, ilab_export_csv_filename)

    ilab_export_csv_file = open(ilab_export_csv_pathname, "w")

    ilab_export_csv_dictwriter = csv.DictWriter(ilab_export_csv_file, fieldnames=ilab_csv_headers)

    ilab_export_csv_dictwriter.writeheader()

    for pi_tag in pi_tag_list:
        print " %s" % pi_tag

        _ = output_ilab_csv_data_for_consulting(ilab_export_csv_dictwriter, pi_tag,
                                                ilab_service_id_consulting_free, ilab_service_id_consulting_paid,
                                                begin_month_timestamp, end_month_timestamp)

    # Close the iLab export CSV file.
    ilab_export_csv_file.close()


#
# Generates the iLab Cluster CSV entries for a particular pi_tag.
#
# It uses dicts pi_tag_to_folder_sizes, pi_tag_to_username_cpus, and pi_tag_to_job_tag_cpus.
#
def output_ilab_csv_data_for_cluster(csv_dictwriter, pi_tag,
                                     storage_service_id, computing_service_id,
                                     begin_month_timestamp, end_month_timestamp):

    # If this PI doesn't have a service request ID, skip them.
    if pi_tag_to_ilab_service_req_id[pi_tag] == '' or pi_tag_to_ilab_service_req_id[pi_tag] == 'N/A':
        print "  Skipping cluster for %s: no iLab service request ID" % (pi_tag)
        return False

    purchased_on_date = from_timestamp_to_date_string(end_month_timestamp-1) # Last date of billing period.

    ###
    #
    # STORAGE Subtable
    #
    ###

    #
    # Get the Service ID for "Local Storage".
    #
    total_storage_sizes = 0.0

    for (folder, size, pctage) in pi_tag_to_folder_sizes[pi_tag]:

        # Note format: <folder> [<pct>%, if not 0%]
        note = "%s" % (folder)

        if 0.0 < pctage < 1.0:
            note += " [%d%%]" % (pctage * 100)

        quantity = size * pctage

        if quantity > 0.0:
            output_ilab_csv_data_row(csv_dictwriter, pi_tag, purchased_on_date, storage_service_id, note, quantity)

            total_storage_sizes += size


    ###
    #
    # COMPUTING Subtable
    #
    ###

    #
    # Get the Service ID for "Local Cluster Computing".
    #
    total_computing_cpuhrs  = 0.0

    # Get the job details for the users associated with this PI.
    if len(pi_tag_to_username_cpus[pi_tag]) > 0:

        for (username, cpu_core_hrs, pctage) in pi_tag_to_username_cpus[pi_tag]:

            fullname = username_to_user_details[username][1]

            # Note format: <user-name> (<user-ID>) [<pct>%, if not 0%]
            note = "User %s (%s)" % (fullname, username)

            if 0.0 < pctage < 1.0:
                note += " [%d%%]" % (pctage * 100)

            quantity = cpu_core_hrs * pctage

            if quantity > 0.0:
                output_ilab_csv_data_row(csv_dictwriter, pi_tag, purchased_on_date, computing_service_id, note, quantity)

                total_computing_cpuhrs += cpu_core_hrs

    else:
        # No users for this PI.
        pass

    # Get the job details for the job tags associated with this PI.
    if len(pi_tag_to_job_tag_cpus[pi_tag]) > 0:

        for (job_tag, cpu_core_hrs, pctage) in pi_tag_to_job_tag_cpus[pi_tag]:

            # Note format: Job Tag <job-tag>) [<pct>%, if not 0%]
            note = "Job Tag %s" % (job_tag)

            if pctage < 1.0:
                note += " [%d%%]" % (pctage * 100)

            quantity = cpu_core_hrs * pctage

            if quantity > 0.0:
                output_ilab_csv_data_row(csv_dictwriter, pi_tag, purchased_on_date, computing_service_id, note, quantity)

                total_computing_cpuhrs += cpu_core_hrs

    else:
        # No job tags for this PI.
        pass

    return True

#
# Generates the iLab Cloud CSV entries for a particular pi_tag.
#
# It uses dicts pi_tag_to_cloud_account_pctages and cloud_project_account_to_cloud_details.
#
def output_ilab_csv_data_for_cloud(csv_dictwriter, pi_tag, cloud_service_id,
                                   begin_month_timestamp, end_month_timestamp):

    # If this PI doesn't have a service request ID, skip them.
    if pi_tag_to_ilab_service_req_id[pi_tag] == '' or pi_tag_to_ilab_service_req_id[pi_tag] == 'N/A':
        print "  Skipping cloud for %s: no iLab service request ID" % (pi_tag)
        return False

    purchased_on_date = from_timestamp_to_date_string(end_month_timestamp-1) # Last date of billing period.

    # Get PI Last name for some situations below.
    (_, pi_last_name, _) = pi_tag_to_names_email[pi_tag]

    # Get list of (account, %ages) tuples for given PI.
    for (account, pctage) in pi_tag_to_cloud_account_pctages[pi_tag]:

        for project in cloud_account_to_cloud_projects[account]:

            # Get list of cloud items to charge PI for.
            cloud_details = cloud_project_account_to_cloud_details[(project, account)]

            for (platform, description, dates, quantity, uom, amount) in cloud_details:

                # If the quantity is given, make a string of it and its unit-of-measure.
                if quantity != '':
                    quantity_str = " @ %s %s" % (quantity, uom)
                else:
                    quantity_str = ''

                if project != '':
                    proj_str = project
                else:
                    proj_str = 'Misc charges/credits for %s' % pi_last_name

                note = "Google :: %s : %s%s" % (proj_str, description, quantity_str)

                if pctage < 1.0:
                    note += " [%d%%]" % (pctage * 100)

                # Calculate the amount to charge the PI based on their percentage.
                pi_amount = amount * pctage

                # Write out the iLab export line.
                output_ilab_csv_data_row(csv_dictwriter, pi_tag, purchased_on_date, cloud_service_id, note, pi_amount)

    return True


def output_ilab_csv_data_for_consulting(csv_dictwriter, pi_tag,
                                        consulting_free_service_id, consulting_paid_service_id,
                                        begin_month_timestamp, end_month_timestamp):

    # If this PI doesn't have a service request ID, skip them.
    if pi_tag_to_ilab_service_req_id[pi_tag] == '' or pi_tag_to_ilab_service_req_id[pi_tag] == 'N/A':
        print "  Skipping consulting for %s: no iLab service request ID" % (pi_tag)
        return False

    for (date, summary, hours, travel_hours, cumul_hours) in consulting_details[pi_tag]:

        date_timestamp = from_excel_date_to_timestamp(date)
        purchased_on_date = from_excel_date_to_date_string(date)

        # If this transaction occurred within this month:
        if begin_month_timestamp <= date_timestamp < end_month_timestamp:

            #
            # Calculate the number of free hours and paid hours in this transaction.
            #
            start_hours_used = cumul_hours - hours - travel_hours

            if start_hours_used < CONSULTING_HOURS_FREE:
                free_hours_remaining = CONSULTING_HOURS_FREE - start_hours_used
            else:
                free_hours_remaining = 0

            if hours < free_hours_remaining:
                free_hours_used = hours
            else:
                free_hours_used = free_hours_remaining

            paid_hours_used = hours - free_hours_used + (travel_hours * CONSULTING_TRAVEL_RATE_DISCOUNT)

            # Write out the iLab export line for the free hours used.
            if free_hours_used > 0:
                output_ilab_csv_data_row(csv_dictwriter, pi_tag, purchased_on_date, consulting_free_service_id,
                                         summary, free_hours_used)

            # Write out the iLab export line for the paid hours used.
            if paid_hours_used > 0:
                output_ilab_csv_data_row(csv_dictwriter, pi_tag, purchased_on_date, consulting_paid_service_id,
                                         summary, paid_hours_used)


def output_ilab_csv_data_row(csv_dictwriter, pi_tag, end_month_string, service_id, note, amount):

    # If this PI doesn't have a service request ID, skip them.
    if pi_tag_to_ilab_service_req_id[pi_tag] == '' or pi_tag_to_ilab_service_req_id[pi_tag] == 'N/A':
        print "  Skipping %s: no iLab service request ID" % (pi_tag)
        return

    # Create a dictionary to be written out as CSV.
    csv_dict = dict()
    csv_dict['owner_email'] = pi_tag_to_names_email[pi_tag][2]
    csv_dict['pi_email']    = ''
    csv_dict['service_request_id'] = int(pi_tag_to_ilab_service_req_id[pi_tag])
    csv_dict['purchased_on'] = end_month_string  # Last date of billing period.
    csv_dict['service_id'] = service_id

    csv_dict['note'] = note
    csv_dict['service_quantity'] = amount

    csv_dictwriter.writerow(csv_dict)


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
parser.add_argument("-g", "--google_invoice_csv",
                    default=None,
                    help="The Google Invoice CSV file")
parser.add_argument("-r", "--billing_root",
                    default=None,
                    help='The Billing Root directory [default = None]')
parser.add_argument("-t", "--ilab_template",
                    default=None,
                    help='The iLab export file template [default = None]')
parser.add_argument("-a", "--ilab_available_services",
                    default=None,
                    help='The iLab available services file [default = None]')
parser.add_argument("-c", "--skip_cluster", action="store_true",
                    default=False,
                    help="Don't output cluster iLab file. [default = False]")
parser.add_argument("-l", "--skip_cloud", action="store_true",
                    default=False,
                    help="Don't output cloud iLab file. [default = False]")
parser.add_argument("-n", "--skip_consulting", action="store_true",
                    default=False,
                    help="Don't output consulting iLab file. [default = False]")
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

# Confirm that BillingDetails file exists.
if not os.path.exists(billing_details_file):
    billing_details_file = None

# If Google Invoice CSV given, use that, else look in BillingRoot.
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

#
# Output the state of arguments.
#
print "GENERATING ILAB EXPORT FOR %02d/%d:" % (month, year)
print "  BillingConfigFile: %s" % (billing_config_file)
print "  BillingRoot: %s" % billing_root
print "  BillingDetailsFile: %s" % (billing_details_file)
print "  GoogleInvoiceCSV: %s" % (google_invoice_csv)
print

#
# Build data structures.
#
print "Building configuration data structures."

# Determine whether we should read in Cloud data from the BillingConfig spreadsheet.
# We should if the BillingConfig spreadsheet has a Cloud sheet.
read_cloud_data = ("Cloud" in billing_config_wkbk.sheet_names())

build_global_data(billing_config_wkbk, begin_month_timestamp, end_month_timestamp, read_cloud_data)

if billing_details_file is not None:
    ###
    #
    # Read the BillingDetails workbook.
    #
    ###

    # Open the BillingDetails workbook.
    print "Opening BillingDetails workbook..."
    billing_details_wkbk = xlrd.open_workbook(billing_details_file)

###
#
# Read in the iLab Export File template, if available.
#
###
if args.ilab_template is not None:

    ilab_template_file = open(args.ilab_template)
    csv_reader = csv.reader(ilab_template_file)
    ilab_csv_headers = csv_reader.next()
    ilab_template_file.close()

else:
    ilab_csv_headers = DEFAULT_CSV_HEADERS

###
#
# Read in the iLab Core Service file, if available.
#
# Set the variables as seen below:
#
###
ilab_service_id_local_computing = None
ilab_service_id_local_storage   = None
ilab_service_id_google_passthrough = None
ilab_service_id_consulting_free = None
ilab_service_id_consulting_paid = None

if args.ilab_available_services is not None:

    ilab_available_services_file = open(args.ilab_available_services)

    csv_dictreader = csv.DictReader(ilab_available_services_file)

    for available_services_row_dict in csv_dictreader:

        # Examine the "Name" column from the available services table.
        row_name_col = available_services_row_dict.get(AVAILABLE_SERVICES_COLUMN_NAME)
        if row_name_col is not None:

            if row_name_col == DEFAULT_AVAILABLE_SERVICES_ID_DICT['Local Computing'][0]:
                ilab_service_id_local_computing = available_services_row_dict[AVAILABLE_SERVICES_COLUMN_SERVICE_ID]
            elif row_name_col == DEFAULT_AVAILABLE_SERVICES_ID_DICT['Local Storage'][0]:
                ilab_service_id_local_storage   = available_services_row_dict[AVAILABLE_SERVICES_COLUMN_SERVICE_ID]
            elif row_name_col == DEFAULT_AVAILABLE_SERVICES_ID_DICT['Cloud Services'][0]:
                ilab_service_id_google_passthrough = available_services_row_dict[AVAILABLE_SERVICES_COLUMN_SERVICE_ID]
            elif row_name_col == DEFAULT_AVAILABLE_SERVICES_ID_DICT['Consulting Free'][0]:
                ilab_service_id_consulting_free = available_services_row_dict[AVAILABLE_SERVICES_COLUMN_SERVICE_ID]
            elif row_name_col == DEFAULT_AVAILABLE_SERVICES_ID_DICT['Consulting Paid'][0]:
                ilab_service_id_consulting_paid = available_services_row_dict[AVAILABLE_SERVICES_COLUMN_SERVICE_ID]

    ilab_available_services_file.close()

    # If we can't find the entries we need from the given available services file,
    #  mention that and exit.
    end_run = False
    if ilab_service_id_local_computing is None:
        print >> sys.stderr, "available services list: No entry for Local Computing"
        end_run = True
    if ilab_service_id_local_storage is None:
        print >> sys.stderr, "available services list: No entry for Local Storage"
        end_run = True
    if ilab_service_id_google_passthrough is None:
        print >> sys.stderr, "available services list: No entry for Cloud Services"
        end_run = True
    if ilab_service_id_consulting_free is None:
        print >> sys.stderr, "available services list: No entry for Consulting Free"
        end_run = True
    if ilab_service_id_consulting_paid is None:
        print >> sys.stderr, "available services list: No entry for Consulting Paid"
        end_run = True

    if end_run:
        print >> sys.stderr, "Problems with available services file: ending run."
        sys.exit(-1)

else:
    # Use default values if no available services file.
    ilab_service_id_local_computing = DEFAULT_AVAILABLE_SERVICES_ID_DICT['Local Computing'][1]
    ilab_service_id_local_storage   = DEFAULT_AVAILABLE_SERVICES_ID_DICT['Local Storage'][1]
    ilab_service_id_google_passthrough = DEFAULT_AVAILABLE_SERVICES_ID_DICT['Cloud Services'][1]
    ilab_service_id_consulting_free = DEFAULT_AVAILABLE_SERVICES_ID_DICT['Consulting Free'][1]
    ilab_service_id_consulting_paid = DEFAULT_AVAILABLE_SERVICES_ID_DICT['Consulting Paid'][1]


#####
#
# Output Cluster data into iLab Cluster export file, if requested.
#
####
if billing_details_file is not None and not args.skip_cluster:
    process_cluster_data()

###
#
# Output Cloud data into iLab Cloud export file, if requested.
#   Read Google Invoice, if given, else use data from BillingDetails file.
#
###
if billing_details_file is not None and not args.skip_cloud:
    process_cloud_data()


#####
#
# Output Consulting data into iLab Cluster export file, if requested.
#
####
if not args.skip_consulting:
    process_consulting_data()