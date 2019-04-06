#!/usr/bin/env python

#===============================================================================
#
# billing_common.py - Set of common utilities and variables to help in the billing scripts.
#
# ARGS:
#
# SWITCHES:x
#
# OUTPUT:
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
import calendar
from collections import OrderedDict
import csv
import datetime
import os

import xlrd
import xlsxwriter


#=====
#
# CONSTANTS
#
#=====

#
# Prefixes for all files created.
#
# Prefix of SGE accounting snapshot file name.
SGEACCOUNTING_PREFIX = "SGEAccounting"
# Prefix of Slurm accounting snapshot file name.
SLURMACCOUNTING_PREFIX = "SlurmAccounting"
# Prefix of the Google Invoice CSV file name.
GOOGLE_INVOICE_PREFIX = "GoogleInvoice"
# Prefix of BillingDetails spreadsheet file name.
BILLING_DETAILS_PREFIX = "BillingDetails"
# Prefix of the BillingNotifs spreadsheets file names.
BILLING_NOTIFS_PREFIX = "GBSCBilling"
# Prefix of the iLab export files.
ILAB_EXPORT_PREFIX = "BillingiLab"
# Prefix of the consulting spreadsheet.
CONSULTING_PREFIX = "BaaSTimesheet"
# Prefix of the storage usage CSV file.
STORAGE_PREFIX = "StorageUsage"

#
# Mapping from BillingConfig sheets to their column headers.
#
BILLING_CONFIG_SHEET_COLUMNS = {
    'PIs'     : ['PI First Name', 'PI Last Name', 'PI Tag', 'Old PI Tag', 'PI Email', 'Group Name', 'PI Folder',
                 'Cluster?', 'Google Cloud?', 'BaaS?', 'Affiliation',
                 'iLab Service Request ID', 'iLab Service Request Name', 'iLab Service Request Owner',
                 'ExPORTER PI ID',
                 'Date Added', 'Date Removed'],
    'Users'   : ['Username', 'Email', 'Full Name', 'PI Tag', '%age', 'Date Added', 'Date Removed'],
    'Folders' : ['Folder', 'PI Tag', '%age', 'Method', 'Date Added', 'Date Removed'],
    'Accounts' : ['Account', 'PI Tag', '%age', 'Date Added', 'Date Removed'],
    'Cloud'   : ['Platform', 'Project', 'Project Number', 'Project ID', 'Account', 'PI Tag', '%age', 'Date Added', 'Date Removed'],
    'Rates'   : ['Type', 'Amount', 'Unit', 'Time'],
    'Config'  : ['Key', 'Value']
}

# Mapping from sheet name to the column headers within that sheet.
BILLING_DETAILS_SHEET_COLUMNS = OrderedDict( (
    ('Storage'   , ('Date Measured', 'Timestamp', 'Folder', 'Size', 'Used')),
    ('Computing' , ('Job Date', 'Job Timestamp', 'Username', 'Job Name', 'Account', 'Node', 'Cores', 'Wallclock Secs', 'Job ID')),
    ('Nonbillable Jobs', ('Job Date', 'Job Timestamp', 'Username', 'Job Name', 'Account', 'Node', 'Cores', 'Wallclock Secs', 'Job ID', 'Reason')),
    ('Failed Jobs', ('Job Date', 'Job Timestamp', 'Username', 'Job Name', 'Account', 'Node', 'Cores', 'Wallclock Secs', 'Job ID', 'Failed Code')),
    ('Cloud', ('Platform', 'Account', 'Project', 'Description', 'Dates', 'Quantity', 'Unit of Measure', 'Charge')),
    ('Consulting', ('Date', 'PI Tag', 'Hours', 'Travel Hours', 'Participants', 'Summary', 'Notes', 'Cumul Hours')) )
)

# Mapping from sheet name to the column headers within that sheet.
BILLING_NOTIFS_SHEET_COLUMNS = OrderedDict( (
    ('Billing',   () ),  # Billing sheet is not columnar.
    ('Lab Users', ('Username', 'Full Name', 'Email', 'Date Added', 'Date Removed') ),
    ('Computing Details' , ('Job Date', 'Username', 'Job Name', 'Job Tag', 'Node', 'CPU-core Hours', 'Job ID', '%age') ),
    ('Cloud Details', ('Platform', 'Project', 'Description', 'Dates', 'Quantity', 'Unit of Measure', 'Charge', '%age', 'Lab Cost') ),
    ('Consulting Details', ('Date', 'Summary', 'Notes', 'Participants', 'Hours', 'Travel Hours', 'Cumul Hours')),
    ('Rates'      , ('Type', 'Amount', 'Unit', 'Time' ) )
) )

# Mapping from sheet name in BillingAggregate workbook to the column headers within that sheet.
BILLING_AGGREG_SHEET_COLUMNS = OrderedDict( [
    ('Totals', ('PI First Name', 'PI Last Name', 'PI Tag', 'iLab Request Name', 'Storage', 'Computing', 'Cloud', 'Consulting', 'Total Charges') )
] )

# OGE accounting file column info:
# http://manpages.ubuntu.com/manpages/lucid/man5/sge_accounting.5.html
ACCOUNTING_FIELDS = (
    'qname', 'hostname', 'group', 'owner', 'job_name', 'job_number',        # Fields 0-5
    'account', 'priority', 'submission_time', 'start_time', 'end_time',     # Fields 6-10
    'failed', 'exit_status', 'ru_wallclock', 'ru_utime', 'ru_stime',        # Fields 11-15
    'ru_maxrss', 'ru_ixrss', 'ru_ismrss', 'ru_idrss', 'ru_isrss', 'ru_minflt', 'ru_majflt',  # Fields 16-22
    'ru_nswap', 'ru_inblock', 'ru_oublock', 'ru_msgsnd', 'ru_msgrcv', 'ru_nsignals',  # Fields 23-28
    'ru_nvcsw', 'ru_nivcsw', 'project', 'department', 'granted_pe', 'slots',  # Fields 29-34
    'task_number', 'cpu', 'mem', 'io', 'category', 'iow', 'pe_taskid', 'max_vmem', 'arid',  # Fields 35-43
    'ar_submission_time'                                                    # Field 44
)

# OGE accounting failed codes which invalidate the accounting entry.
# From https://arc.liv.ac.uk/SGE/htmlman/htmlman5/sge_status.html
ACCOUNTING_FAILED_CODES = (1,3,4,5,6,7,8,9,10,11,18,19,20,21,26,27,28,29,36,38)

# List of hostname prefixes to determine which jobs on which nodes are billable.
BILLABLE_HOSTNAME_PREFIXES = ['scg1', 'scg3-1', 'scg3-2', 'scg4',
                              'sgisummit-rcf-111', 'sgisummit-frcf-111', # WAS scg3-2
                              'sgiuv20-rcf-111',                         # WAS scg3-1-fatnode
                              'dper730xd-srcf-d16',                      # WAS scg4-h17
                              'dper930-srcf-d15-05',                     # WAS scg4-h16-05
                              'dper7425-srcf-d15'                        # Nodes installed 9/2018.
                              ]
NONBILLABLE_HOSTNAME_PREFIXES = ['scg3-0',
                                 'dper910-rcf-412-20', 'greenie',        # Synonyms for greenie
                                 'hppsl230s-rcf-412',                    # WAS scg3-0
                                 'sgiuv300-srcf',                        # The supercomputer
                                 'cfxs2600gz-rcf-114',                   # Data Mover nodes
                                 'dper7425-srcf-d10-37',                 # Free node from Dell
                                 'dper740xd-srcf-d6-22',                 # PI server: Khavari
                                 'dper740xd-srcf-d5-35',                 # PI server: Howard Chang
                                 'None assigned'
                                 ]

# Job tag/account prefixes for PI Tags. [Format: <Prefix>_<PI_TAG>]
ACCOUNT_PREFIXES = ['apps', ''baas', 'baas_lab', 'baas_prj', 'nih', 'prj']
# List of accounts to ignore.
IGNORED_ACCOUNTS = ['large_mem', 'default']

# Beginning of billing process.
# 8/31/13 00:00:00 GMT (one day before 9/1/13, to represent things that existed before billing started).
EARLIEST_VALID_DATE_EXCELDATE = 41517.0

# The maximum number of rows in any one Excel sheet.
EXCEL_MAX_ROWS = 1048576

# Top-level directories for various file systems.
GPFS_TOPLEVEL_DIRECTORIES = ['/srv/gsfs0']
ISILON_TOPLEVEL_DIRECTORIES = ['/ifs', '/BaaS', '/labs', '/projects', '/reference', '/scg' ]

# Commands for determining folder quotas and usages.
QUOTA_EXECUTABLE_GPFS = ['ssh', 'root@scg-gs0', '/usr/lpp/mmfs/bin/mmlsquota', '-j']
QUOTA_EXECUTABLE_ISILON = ['df']
USAGE_EXECUTABLE = ['du', '-s']
STORAGE_BLOCK_SIZE_ARG = ['--block-size=1G']  # Works in all above commands.

# Pathname to root of PI project directories.
PI_PROJECT_ROOT_DIR = '/labs'

# How many hours of consulting are free.
CONSULTING_HOURS_FREE = 1

# What is the discount rate for travel hours?
CONSULTING_TRAVEL_RATE_DISCOUNT = 0.5


#=====
#
# FUNCTIONS
#
#=====

# This method takes in an xlrd Sheet object and a column name,
# and returns all the values from that column headed by that name.
def sheet_get_named_column(sheet, col_name):

    header_row = sheet.row_values(0)

    for idx in range(len(header_row)):
        if header_row[idx] == col_name:
           col_name_idx = idx
           break
    else:
        return None

    return sheet.col_values(col_name_idx,start_rowx=1)

# This function returns the dict of values in a BillingConfig's Config sheet.
def config_sheet_get_dict(wkbk):

    config_sheet = wkbk.sheet_by_name("Config")

    config_keys   = sheet_get_named_column(config_sheet, "Key")
    config_values = sheet_get_named_column(config_sheet, "Value")

    return dict(zip(config_keys, config_values))


# Read the Config sheet of the BillingConfig workbook to
# get the BillingRoot directory and the SGEAccountingFile.
# Returns a tuple of (BillingRoot, SGEAccountingFile).
def read_config_sheet(wkbk):

    config_dict = config_sheet_get_dict(wkbk)

    accounting_file = config_dict.get("SGEAccountingFile")
    billing_root    = config_dict.get("BillingRoot", os.getcwd())

    return (billing_root, accounting_file)


#
# This suite of functions converts to/from timestamps, Excel dates, and date strings.
#
def from_timestamp_to_excel_date(timestamp):
    return timestamp/86400.0 + 25569
def from_excel_date_to_timestamp(excel_date):
    return int((excel_date - 25569) * 86400.0)
def from_timestamp_to_date_string(timestamp):
    return datetime.datetime.utcfromtimestamp(timestamp).strftime("%m/%d/%Y")
def from_excel_date_to_date_string(excel_date):
    return from_timestamp_to_date_string(from_excel_date_to_timestamp(excel_date))

def from_ymd_date_to_timestamp(year, month, day):
    return int(calendar.timegm(datetime.date(year, month, day).timetuple()))
def from_date_string_to_timestamp(date_str):
    return int(calendar.timegm(datetime.datetime.strptime(date_str, "%m/%d/%y").timetuple()))

#
# This function removes the Unicode characters from a string.
#
def remove_unicode_chars(s):
    return "".join(i for i in s if ord(i)<128)


# Filters a list of lists using a parallel list of [date_added, date_removed]'s.
# Returns the elements in the first list which are valid with the month date range given.
def filter_by_dates(obj_list, date_list, begin_month_exceldate, end_month_exceldate):

    output_list = []

    for (obj, (date_added, date_removed)) in zip(obj_list, date_list):

        # If the date added is BEFORE the end of this month, and
        #    the date removed is AFTER the beginning of this month,
        # then save the account information in the mappings.
        if date_added < end_month_exceldate and (date_removed == '' or date_removed >= begin_month_exceldate):
            output_list.append(obj)

    return output_list
