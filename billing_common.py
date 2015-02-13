#!/usr/bin/env python

#===============================================================================
#
# billing_common.py - Set of common utilities and variables to help in the billing scripts.
#
# ARGS:
#
# SWITCHES:
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
# Prefix of BillingDetails spreadsheet file name.
BILLING_DETAILS_PREFIX = "BillingDetails"
# Prefix of the BillingNotifs spreadsheets file names.
BILLING_NOTIFS_PREFIX = "GBSCBilling"

#
# Mapping from BillingConfig sheets to their column headers.
#
BILLING_CONFIG_SHEET_COLUMNS = {
    'Rates'   : ['Type', 'Amount', 'Unit', 'Time'],
    'PIs'     : ['PI First Name', 'PI Last Name', 'PI Tag', 'Group Name', 'PI Email', 'iLab Service Request ID', 'Date Added', 'Date Removed'],
    'Folders' : ['Folder', 'PI Tag', '%age', 'Method', 'Date Added', 'Date Removed'],
    'Users'   : ['PI Tag', 'Username', 'Email', 'Full Name', 'Date Added', 'Date Removed'],
    'JobTags' : ['Job Tag', 'PI Tag', '%age', 'Date Added', 'Date Removed'],
    'Config'  : ['Key', 'Value']
}

# Mapping from sheet name to the column headers within that sheet.
BILLING_DETAILS_SHEET_COLUMNS = OrderedDict((
    ('Storage'   , ('Date Measured', 'Timestamp', 'Folder', 'Size', 'Used')),
    ('Computing' , ('Job Date', 'Job Timestamp', 'Username', 'Job Name', 'Account', 'Node', 'Cores', 'Wallclock Secs', 'Job ID')),
#    ('Consulting', ('Work Date', 'Item', 'Hours', 'PI')),
    ('Nonbillable Jobs', ('Job Date', 'Job Timestamp', 'Username', 'Job Name', 'Account', 'Node', 'Cores', 'Wallclock Secs', 'Job ID', 'Reason')),
    ('Failed Jobs', ('Job Date', 'Job Timestamp', 'Username', 'Job Name', 'Account', 'Node', 'Cores', 'Wallclock Secs', 'Job ID', 'Failed Code'))
) )

# Mapping from sheet name to the column headers within that sheet.
BILLING_NOTIFS_SHEET_COLUMNS = OrderedDict( (
    ('Billing'    , () ),  # Billing sheet is not columnar.
    ('Lab Users'  , ('Username', 'Full Name', 'Email', 'Date Added', 'Date Removed') ),
    ('Computing Details' , ('Job Date', 'Username', 'Job Name', 'Job Tag', 'CPU-core Hours', 'Job ID', '%age') ),
    ('Rates'      , ('Type', 'Amount', 'Unit', 'Time' ) )
) )

# Mapping from sheet name in BillingAggregate workbook to the column headers within that sheet.
BILLING_AGGREG_SHEET_COLUMNS = OrderedDict( [
    #('Totals', ('PI First Name', 'PI Last Name', 'PI Tag', 'Storage', 'Computing', 'Consulting', 'Total Charges') )
    ('Totals', ('PI First Name', 'PI Last Name', 'PI Tag', 'Storage', 'Computing', 'Total Charges') )
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

# List of hostname prefixes to use for billing purposes.
BILLABLE_HOSTNAME_PREFIXES = ['scg1', 'scg3-1']

# Beginning of billing process.
# 8/31/13 00:00:00 GMT (one day before 9/1/13, to represent things that existed before billing started).
EARLIEST_VALID_DATE_EXCELDATE = 41517.0

# Pathname to root of PI project directories.
PI_PROJECT_ROOT_DIR = '/srv/gsfs0/projects'

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
    return (excel_date - 25569) * 86400.0
def from_timestamp_to_date_string(timestamp):
    return datetime.datetime.utcfromtimestamp(timestamp).strftime("%m/%d/%Y")
def from_excel_date_to_date_string(excel_date):
    return from_timestamp_to_date_string(from_excel_date_to_timestamp(excel_date))

def from_ymd_date_to_timestamp(year, month, day):
    return int(calendar.timegm(datetime.date(year, month, day).timetuple()))

#
# This function removes the Unicode characters from a string.
#
def remove_unicode_chars(s):
    return "".join(i for i in s if ord(i)<128)