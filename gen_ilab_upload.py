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
from collections import defaultdict
import csv
import datetime
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
global BILLING_NOTIFS_PREFIX

# Prefix for the iLab Export CSV filename.
ILAB_EXPORT_PREFIX = 'BillingiLab'

# Default headers for the ilab Export CSV file (if not read in from iLab template file).
DEFAULT_CSV_HEADERS = ['service_id','note','service_quantity','purchased_on',
                       'service_request_id','owner_email','pi_email']

# Default service IDs (if not read in from iLab Core Services file).
DEFAULT_SERVICE_ID_LOCAL_STORAGE   = 1991
DEFAULT_SERVICE_ID_LOCAL_COMPUTING = 1992
DEFAULT_SERVICE_ID_GOOGLE_STORAGE  = 2191
DEFAULT_SERVICE_ID_GOOGLE_EGRESS   = 2192

# Service ID names (for reading in service IDs from iLab Core Services file).
CORE_SERVICES_COLUMN_NAME = 'Name'
CORE_SERVICES_COLUMN_SERVICE_ID = 'Service ID'

CORE_SERVICES_NAME_LOCAL_STORAGE = 'Local Storage'
CORE_SERVICES_NAME_LOCAL_COMPUTING = 'Local Cluster Computing'

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

# Mapping from pi_tag to list of [date, username, job_name, account, cpu_core_hrs, jobID, %age].
pi_tag_to_sge_job_details = defaultdict(list)

# Mapping from pi_tag to list of [username, date_added, date_removed, %age].
pi_tag_to_user_details = defaultdict(list)


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

    pi_tag_to_ilab_service_req_id = dict(zip(pi_tag_list,pi_ilab_ids))

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

    job_tag_rows = zip(job_tags, pi_tags, pctages)

    for (job_tag, pi_tag, pctage) in job_tag_rows:
        job_tag_to_pi_tag_pctages[job_tag].append([pi_tag, pctage])

    #
    # Create mapping from folder to list of pi_tags and %ages.
    #
    global folder_to_pi_tag_pctages

    folders = sheet_get_named_column(folders_sheet, "Folder")
    pi_tags = sheet_get_named_column(folders_sheet, "PI Tag")
    pctages = sheet_get_named_column(folders_sheet, "%age")

    folder_rows = zip(folders, pi_tags, pctages)

    for (folder, pi_tag, pctage) in folder_rows:
        folder_to_pi_tag_pctages[folder].append([pi_tag, pctage])


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
# pi_tag_to_sge_job_details dicts.
def read_computing_sheet(wkbk):

    global pi_tag_to_sge_job_details
    global pi_tag_to_job_tag_cpus
    global pi_tag_to_username_cpus

    computing_sheet = wkbk.sheet_by_name("Computing")

    for row in range(1,computing_sheet.nrows):

        (job_date, job_timestamp, job_username, job_name, account, node, cores, wallclock, jobID) = \
            computing_sheet.row_values(row)

        # Calculate CPU-core-hrs for job.
        cpu_core_hrs = cores * wallclock / 3600.0  # wallclock is in seconds.

        # Rename this variable for easier understanding.
        job_tag = account

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

                        # Increment the user's CPUs
                        if username == job_username:
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


#
# Generates the iLab CSV entries for a particular pi_tag.
#
# It uses dicts pi_tag_to_folder_sizes, pi_tag_to_username_cpus, and pi_tag_to_job_tag_cpus.
#
def generate_ilab_csv_file(csv_dictwriter, pi_tag,
                           storage_service_id, computing_service_id,
                           begin_month_timestamp, end_month_timestamp):

    # If this PI doesn't have a service request ID, skip them.
    if pi_tag_to_ilab_service_req_id[pi_tag] == '':
        print "  Skipping %s: no service request ID" % (pi_tag)
        return

    # Create a dictionary to be written out as CSV.
    csv_dict = dict()
    csv_dict['owner_email'] = pi_tag_to_names_email[pi_tag][2]
    csv_dict['pi_email']    = ''
    csv_dict['service_request_id'] = pi_tag_to_ilab_service_req_id[pi_tag]
    csv_dict['purchased_on'] = from_timestamp_to_date_string(end_month_timestamp-1) # Last date of billing period.

    ###
    #
    # STORAGE Subtable
    #
    ###

    #
    # Get the Service ID for "Local Storage".
    #
    csv_dict['service_id'] = storage_service_id

    total_storage_sizes = 0.0

    for (folder, size, pctage) in pi_tag_to_folder_sizes[pi_tag]:

        # Note format: <folder> [<pct>%, if not 0%]
        note = "%s" % (folder)

        if 0.0 < pctage < 1.0:
            note += " [%d%%]" % (pctage * 100)

        quantity = size * pctage

        csv_dict['note'] = note
        csv_dict['service_quantity'] = quantity

        if quantity > 0.0:
            csv_dictwriter.writerow(csv_dict)

            total_storage_sizes += size


    ###
    #
    # COMPUTING Subtable
    #
    ###

    #
    # Get the Service ID for "Local Cluster Computing".
    #
    csv_dict['service_id'] = computing_service_id

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

            csv_dict['note'] = note
            csv_dict['service_quantity'] = quantity

            if quantity > 0.0:
                csv_dictwriter.writerow(csv_dict)

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
                note += " [%d%%]" % (pctage)

            quantity = cpu_core_hrs * pctage

            csv_dict['note'] = note
            csv_dict['service_quantity'] = quantity

            if quantity > 0.0:
                csv_dictwriter.writerow(csv_dict)

                total_computing_cpuhrs += cpu_core_hrs

    else:
        # No job tags for this PI.
        pass



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
parser.add_argument("-t", "--ilab_template",
                    default=None,
                    help='The iLab export file template [default = None]')
parser.add_argument("-a", "--ilab_available_services",
                    default=None,
                    help='The iLab available services file [default = None]')
parser.add_argument("-p", "--pi_files", action="store_true",
                    default=False,
                    help='Output PI-specific CSV files [default = False]')
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

###
#
# Read the BillingConfig workbook and build input data structures.
#
###

billing_config_wkbk = xlrd.open_workbook(args.billing_config_file)

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

# Within BillingRoot, create YEAR/MONTH dirs if necessary.
year_month_dir = os.path.join(billing_root, str(year), "%02d" % month)
if not os.path.exists(year_month_dir):
    os.makedirs(year_month_dir)

# If BillingDetails file given, use that, else look in BillingRoot.
if args.billing_details_file is not None:
    billing_details_file = args.billing_details_file
else:
    billing_details_file = os.path.join(year_month_dir, "BillingDetails.%s-%02d.xlsx" % (year, month))

#
# Output the state of arguments.
#
print "GENERATING ILAB EXPORT FOR %02d/%d:" % (month, year)
print "  BillingConfigFile: %s" % (args.billing_config_file)
print "  BillingRoot: %s" % billing_root
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

###
#
# Read in the iLab File template, if available.
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
# Set the variables:
#   ilab_service_id_local_computing
#   ilab_service_id_local_storage
#
###
if args.ilab_available_services is not None:

    ilab_available_services_file = open(args.ilab_available_services)

    csv_dictreader = csv.DictReader(ilab_available_services_file)

    for available_services_row_dict in csv_dictreader:

        row_name_col = available_services_row_dict.get(CORE_SERVICES_COLUMN_NAME)
        if row_name_col is not None:

            if row_name_col == CORE_SERVICES_NAME_LOCAL_COMPUTING:
                ilab_service_id_local_computing = available_services_row_dict[CORE_SERVICES_COLUMN_SERVICE_ID]
            elif row_name_col == CORE_SERVICES_NAME_LOCAL_STORAGE:
                ilab_service_id_local_storage   = available_services_row_dict[CORE_SERVICES_COLUMN_SERVICE_ID]

    ilab_available_services_file.close()

else:
    ilab_service_id_local_computing = DEFAULT_SERVICE_ID_LOCAL_COMPUTING
    ilab_service_id_local_storage   = DEFAULT_SERVICE_ID_LOCAL_STORAGE


###
#
# Open the iLab CSV file for writing out.
#
###

ilab_csv_filename = "%s.%s-%02d.csv" % (ILAB_EXPORT_PREFIX, year, month)
ilab_csv_pathname = os.path.join(year_month_dir, ilab_csv_filename)

csv_file = open(ilab_csv_pathname, "w")

csv_dictwriter = csv.DictWriter(csv_file, fieldnames=ilab_csv_headers)

###
#
# Write iLab export CSV file from output data structures.
#
###

csv_dictwriter.writeheader()

print "Writing iLab export CSV file:"
for pi_tag in sorted(pi_tag_list):

    print " %s" % pi_tag

    generate_ilab_csv_file(csv_dictwriter, pi_tag,
                           ilab_service_id_local_storage, ilab_service_id_local_computing,
                           begin_month_timestamp, end_month_timestamp)

csv_file.close()

