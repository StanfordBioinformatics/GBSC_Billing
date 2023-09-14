#!/usr/bin/env python3

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
import argparse
import calendar
import sys
from collections import OrderedDict
import datetime
import os

from slurm_job_accounting_entry import SlurmJobAccountingEntry

#=====
#
# CONSTANTS
#
#=====

# Directory names for hierarchy under each month
SUBDIR_RAWDATA  = "RawData"
SUBDIR_EXPORTS  = "Exports"
SUBDIR_INVOICES = "Invoices"

#
# Prefixes for all files created.
#

# Prefix for BillingConfig file name.
BILLING_CONFIG_PREFIX = "GBSCBilling_Config"
# Prefix of BillingDetails spreadsheet file name.
BILLING_DETAILS_PREFIX = "GBSCBilling_Details"
# Prefix for the BillingAggregate file name.
BILLING_AGGREGATE_PREFIX = "GBSCBilling_Summary"

## "Invoices" files
# Prefix of the BillingNotifs spreadsheets file names.
BILLING_NOTIFS_PREFIX = "GBSCBilling"

## "Exports" files
# Prefix of the iLab export files.
ILAB_EXPORT_PREFIX = "GBSCBillingiLab"

## "RawData" files
# Prefix of SGE accounting snapshot file name.
SGEACCOUNTING_PREFIX = "SGEAccounting"
# Prefix of Slurm accounting snapshot file name.
SLURMACCOUNTING_PREFIX = "SlurmAccounting"
# Prefix of the Google Invoice CSV file name.
GOOGLE_INVOICE_PREFIX = "GoogleInvoice"
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
    ('Storage'   , ('Date Measured', 'Timestamp', 'Folder', 'Size', 'Used', 'Inodes Quota', 'Inodes Used')),
    ('Computing' , ('Job Date', 'Job Timestamp', 'Username', 'Job Name', 'Account', 'Node', 'Cores', 'Wallclock Secs', 'Job ID')),
    ('Nonbillable Jobs', ('Job Date', 'Job Timestamp', 'Username', 'Job Name', 'Account', 'Node', 'Cores', 'Wallclock Secs', 'Job ID', 'Reason')),
    ('Failed Jobs', ('Job Date', 'Job Timestamp', 'Username', 'Job Name', 'Account', 'Node', 'Cores', 'Wallclock Secs', 'Job ID', 'Failed Code')),
    ('Cloud', ('Platform', 'Account', 'Project', 'Description', 'Dates', 'Quantity', 'Unit of Measure', 'Charge')),
    ('Consulting', ('Date', 'PI Tag', 'Hours', 'Travel Hours', 'Participants', 'Clients', 'Summary', 'Notes', 'Cumul Hours')) )
)

# Mapping from sheet name to the column headers within that sheet.
BILLING_NOTIFS_SHEET_COLUMNS = OrderedDict( (
    ('Billing',   () ),  # Billing sheet is not columnar.
    ('Lab Users', ('Username', 'Full Name', 'Email', 'Date Added', 'Date Removed') ),
    ('Computing Details' , ('Job Date', 'Username', 'Job Name', 'Job Tag', 'Node', 'CPU-core Hours', 'Job ID', '%age') ),
    ('Cloud Details', ('Platform', 'Project', 'Description', 'Dates', 'Quantity', 'Unit of Measure', 'Charge', '%age', 'Lab Cost') ),
    ('Consulting Details', ('Date', 'Summary', 'Notes', 'Participants', 'Clients', 'Hours', 'Travel Hours', 'Cumul Hours')),
    ('Rates'      , ('Type', 'Amount', 'Unit', 'Time', 'iLab Service ID' ) )
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

# SLURM: Which delimiter is used in the accounting files?
SLURMACCOUNTING_DELIMITER = SlurmJobAccountingEntry.DELIMITER_HASH

# List of hostname prefixes to determine which jobs on which nodes are billable.
BILLABLE_HOSTNAME_PREFIXES = ['scg1', 'scg3-1', 'scg3-2', 'scg4',
                              'sgisummit-rcf-111', 'sgisummit-frcf-111', # WAS scg3-2
                              'sgiuv20-rcf-111',                         # WAS scg3-1-fatnode
                              'dper730xd-srcf-d16',                      # WAS scg4-h17
                              'dper930-srcf-d15-05',                     # WAS scg4-h16-05
                              'dper7425-srcf-d15',                       # Nodes installed 9/2018.
                              'cfx22885s-srcf-d11',                      # Nodes installed 3/2021.
                              'cfx4860s-srcf-d11','cfx1265s-srcf-d11-27' # Nodes installed 4/2023.
                              ]

NONBILLABLE_HOSTNAME_PREFIXES = ['scg3-0',
                                 'dper910-rcf-412-20', 'greenie',        # Synonyms for greenie
                                 'hppsl230s-rcf-412',                    # WAS scg3-0
                                 'sgiuv300-srcf',                        # The supercomputer
                                 'cfxs2600gz-rcf-114',                   # Data Mover nodes
                                 'dper7425-srcf-d10-37',                 # Free node from Dell
                                 'smsh11dsu-srcf-d15',                   # Login nodes installed 9/19.
                                 'dper740xd-srcf-d6-22',                 # PI server: Khavari
                                 'dper740xd-srcf-d5-35',                 # PI server: Howard Chang
                                 'smsx10srw-srcf-d15',                   # Login nodes
                                 'smsh11dsu-frcf-212',                   # GSSC-owned nodes
                                 'dper630-frcf-212',                     # Test login node/warm spares for maintenance nodes
                                 'smsx11dsc-frcf-212',                   # Storage node
                                 'ddnr620-frcf-213',                     # Old DDN storage nodes, repurposed
                                 'dper820-frcf-212',                     # Obsolete Sherlock nodes
                                 'smsh11dsu-srcf-d10',                   # Slurm management nodes
                                 'dper7525-srcf-d11-21',                 # OnDemand dev node
                                 'None assigned'
                                 ]

# Job tag/account prefixes for PI Tags. [Format: <Prefix>_<PI_TAG>]
ACCOUNT_PREFIXES = ['apps', 'baas', 'baas_lab', 'baas_prj', 'nih', 'owner', 'org', 'prj']
# List of accounts to ignore.
IGNORED_ACCOUNTS = ['large_mem', 'default']

# Beginning of billing process.
# 8/31/13 00:00:00 GMT (one day before 9/1/13, to represent things that existed before billing started).
EARLIEST_VALID_DATE_EXCELDATE = 41517.0

# The maximum number of rows in any one Excel sheet.
EXCEL_MAX_ROWS = 1048576

# Top-level directories for various file systems.
GPFS_TOPLEVEL_DIRECTORIES = ['/srv/gsfs0']
ISILON_TOPLEVEL_DIRECTORIES = ['/ifs', '/BaaS', '/labs', '/projects', '/reference', '/scg']
TOPLEVEL_DIRECTORIES = ['/BaaS', '/labs', '/projects', '/reference', '/scg']

# Commands for determining folder quotas and usages.
QUOTA_EXECUTABLE_GPFS = ['ssh', 'root@scg-gs0', '/usr/lpp/mmfs/bin/mmlsquota', '-j']
QUOTA_EXECUTABLE_ISILON = ['df']
QUOTA_EXECUTABLE = ['df']
USAGE_EXECUTABLE = ['du', '-s']
INODES_EXECUTABLE = ['df', '-i']
STORAGE_BLOCK_SIZE_ARG = ['--block-size=1G']  # Works in all above commands.

# Pathname to root of PI project directories.
PI_PROJECT_ROOT_DIR = '/labs'

# Subdirectory name for BaaS, found in PI Folders.
BAAS_SUBDIR_NAME = "BaaS"

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

    # header_row = sheet.row_values(0)
    header_row = sheet[1]

    for idx in range(len(header_row)):
        # if header_row[idx] == col_name:
        if header_row[idx].value == col_name:
            col_name_idx = idx
            break
    else:
        return None

    max_row = sheet.max_row

    # return sheet.col_values(col_name_idx,start_rowx=1)
    return list(list(sheet.iter_cols(min_col=col_name_idx+1,max_col=col_name_idx+1,min_row=2,values_only=True))[0])

# This function returns the dict of values in a BillingConfig's Config sheet.
def config_sheet_get_dict(wkbk):

    #config_sheet = wkbk.sheet_by_name("Config")
    config_sheet = wkbk["Config"]

    config_keys   = sheet_get_named_column(config_sheet, "Key")
    config_values = sheet_get_named_column(config_sheet, "Value")

    return dict(list(zip(config_keys, config_values)))


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

# Now adding conversion functions to/from datetime for openpyxl
def from_timestamp_to_datetime(timestamp):
    return datetime.datetime.utcfromtimestamp(timestamp)
def from_datetime_to_timestamp(dt):
    return dt.timestamp()
def from_excel_date_to_datetime(excel_date):
    return from_timestamp_to_datetime(from_excel_date_to_timestamp(excel_date))
def from_datetime_to_excel_date(dt):
    return from_timestamp_to_excel_date(from_datetime_to_timestamp(dt))
def from_ymd_date_to_datetime(year, month, day):
    return from_timestamp_to_datetime(from_ymd_date_to_timestamp(year, month, day))
def from_datetime_to_date_string(dt):
    return dt.strftime("%m/%d/%Y")

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

        if date_added is None: continue

        # If the date added is BEFORE the end of this month, and
        #    the date removed is AFTER the beginning of this month,
        # then save the account information in the mappings.
        if (date_added < end_month_exceldate and
                (date_removed is None or date_removed == '' or date_removed >= begin_month_exceldate)):
            output_list.append(obj)

    return output_list


# A parent parser for arguments which are common across many scripts
def argparse_get_parent_parser():
    parser = argparse.ArgumentParser(add_help=False)

    parser.add_argument("-b", "--billing_config_file",
                        help='The BillingConfig file')
    parser.add_argument("-r", "--billing_root",
                        default=None,
                        help='The Billing Root directory [default = None]')
    parser.add_argument("-y", "--year", type=int, choices=list(range(2013, 2031)),
                        default=None,
                        help="The year to be filtered out. [default = this year]")
    parser.add_argument("-m", "--month", type=int, choices=list(range(1, 13)),
                        default=None,
                        help="The month to be filtered out. [default = last month]")
    parser.add_argument("-v", "--verbose", action="store_true",
                        default=False,
                        help='Get real chatty [default = false]')
    parser.add_argument("-d", "--debug", action="store_true",
                        default=False,
                        help='Get REAL chatty [default = false]')

    return parser

# Assumption: "args" has "year" and "month" fields
def argparse_get_year_month(args):

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
    end_month_timestamp = from_ymd_date_to_timestamp(next_month_year, next_month, 1)

    return year, month, begin_month_timestamp, end_month_timestamp

# Assumption: "args" has "billing_root" and "billing_config_file" fields.
def argparse_get_billingroot_billingconfig(args, year, month):

    # Try to get values from parser.
    billing_config_file = args.billing_config_file
    billing_root        = args.billing_root

    if billing_config_file is None:
        if billing_root is None:
            billing_root = os.getcwd()

        # Find BillingConfig within BillingRoot
        input_subdir = get_subdirectory(billing_root, year, month, "")

        billing_config_file = os.path.join(input_subdir, "{0}.{1:d}-{2:02d}.xlsx".format(BILLING_CONFIG_PREFIX, year, month))
        # Does this file exist?
        if not os.path.exists(billing_config_file):
            print("ArgParse: BillingConfig file {} does not exist".format(billing_config_file),file=sys.stderr)
            sys.exit(-1)

    elif billing_root is None:

        if not os.path.exists(billing_config_file):
            print("ArgParse: BillingConfig file {} does not exist".format(billing_config_file), file=sys.stderr)
            sys.exit(-1)

        # Get the location of the BillingRoot directory from the Config sheet of the BillingConfig workbook.
        billing_config_wkbk = openpyxl.load_workbook(billing_config_file)  # , read_only=True)

        #  Ignore the accounting file from this sheet.
        (billing_root, _) = read_config_sheet(billing_config_wkbk)
        if billing_root is None:
            billing_root = os.getcwd()
        else:
            # Does BillingRoot exist?
            if not os.path.exists(billing_root):
                print("ArgParse: BillingRoot dir {} does not exist".format(billing_root), file=sys.stderr)
                sys.exit(-1)

        billing_config_wkbk.close()

    else:
        # We have both BillingRoot and BillingConfig: do they exist?
        if not os.path.exists(billing_root):
            print("ArgParse: BillingRoot dir {} does not exist".format(billing_root), file=sys.stderr)
            sys.exit(-1)
        if not os.path.exists(billing_config_file):
            print("ArgParse: BillingConfig file {} does not exist".format(billing_config_file), file=sys.stderr)
            sys.exit(-1)

    # Get absolute path for billing_root directory.
    billing_root        = os.path.abspath(billing_root)
    # Get absolute path for billing_config_file.
    billing_config_file = os.path.abspath(billing_config_file)

    return billing_root, billing_config_file


# This function returns an integer for the Fiscal Year given month and year
def get_fiscal_year(year, month):
    if month >= 9:
        return year + 1
    else:
        return year


# This function generates paths below the BillingRoot, creating them if requested
def get_subdirectory(billing_root, year, month, subdir, create_if_nec=False):

    fiscal_year = get_fiscal_year(year, month)

    full_subdir = os.path.join(billing_root, "FY%d" % fiscal_year, str(year), "%02d" % month, subdir)

    if not os.path.exists(full_subdir):
        if create_if_nec:
            os.makedirs(full_subdir)
        else:
            print("get_subdirectory: Can't find %s" % full_subdir, file=sys.stderr)
            return None

    return full_subdir