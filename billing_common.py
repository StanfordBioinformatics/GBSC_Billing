#!/usr/bin/env python

#===============================================================================
#
# billing_common.py - Set of utilities to help in using workbooks
#                   from xlrd and Xlsxwriter.
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
from collections import OrderedDict
import argparse
import os
import os.path
import sys

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
    'PIs'     : ['PI First Name', 'PI Last Name', 'PI Tag', 'Group Name', 'PI Email', 'Date Added'],
    'Folders' : ['Folder', 'PI Tag', '%age', 'By Quota?', 'Date Added'],
    'Users'   : ['PI Tag', 'Username', 'Email', 'Full Name', 'Date Added', 'Date Removed'],
    'JobTags' : ['Job Tag', 'PI Tag', '%age', 'Date Added'],
    'Config'  : ['Key', 'Value']
}

# Mapping from sheet name to the column headers within that sheet.
BILLING_DETAILS_SHEET_COLUMNS = OrderedDict((
    ('Storage'   , ('Date Measured', 'Timestamp', 'Folder', 'Size', 'Used')),
    ('Computing' , ('Job Date', 'Job Timestamp', 'Username', 'Job Name', 'Account', 'Node', 'Cores', 'Wallclock', 'Job ID')),
    ('Consulting', ('Work Date', 'Item', 'Hours', 'PI')),
    ('Nonbillable Jobs', ('Job Date', 'Job Timestamp', 'Username', 'Job Name', 'Account', 'Node', 'Cores', 'Wallclock', 'Job ID', 'Reason')),
    ('Failed Jobs', ('Job Date', 'Job Timestamp', 'Username', 'Job Name', 'Account', 'Node', 'Cores', 'Wallclock', 'Job ID', 'Failed Code'))
) )

# Mapping from sheet name to the column headers within that sheet.
BILLING_NOTIFS_SHEET_COLUMNS = OrderedDict( (
    ('Billing'    , () ),  # Billing sheet is not columnar.
    ('Lab Users'  , ('Username', 'Email', 'Full Name', 'Date Added', 'Date Removed') ),
    ('Computing Details' , ('Job Date', 'Username', 'Job Name', 'Job Tag', 'CPU-core Hours', 'Job ID', '%age') ),
    ('Rates'      , ('Type', 'Amount', 'Unit', 'Time' ) )
) )

# Mapping from sheet name in BillingAggregate workbook to the column headers within that sheet.
BILLING_AGGREG_SHEET_COLUMNS = OrderedDict( [
    ('Totals', ('PI Tag', 'Storage', 'Computing', 'Consulting', 'Total Charges') )
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
    'task_number', 'cpu', 'mem', 'category', 'iow', 'pe_taskid', 'max_vmem', 'arid',  # Fields 35-42
    'ar_submission_time'                                                    # Field 43
)

# OGE accounting failed codes which invalidate the accounting entry.
# From http://docs.oracle.com/cd/E19080-01/n1.grid.eng6/817-6117/chp11-1/index.html
ACCOUNTING_FAILED_CODES = (1,3,4,5,6,7,8,9,10,11,26,27,28)

# List of hostname prefixes to use for billing purposes.
BILLABLE_HOSTNAME_PREFIXES = ['scg1']




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


def config_sheet_get_dict(wkbk):

    config_sheet = wkbk.sheet_by_name("Config")

    config_keys   = sheet_get_named_column(config_sheet, "Key")
    config_values = sheet_get_named_column(config_sheet, "Value")

    return dict(zip(config_keys, config_values))


# Read the Config sheet of the BillingConfig workbook to
# get the BillingRoot directory.  Returns it if present,
# the current directory if not.
def read_config_sheet(wkbk):

    config_dict = config_sheet_get_dict(wkbk)

    accounting_file = config_dict.get("SGEAccountingFile")
    billing_root    = config_dict.get("BillingRoot", os.getcwd())

    return (billing_root, accounting_file)


def from_timestamp_to_excel_date(timestamp):
    return timestamp/86400 + 25569
def from_excel_date_to_timestamp(excel_date):
    return (excel_date - 25569) * 86400
