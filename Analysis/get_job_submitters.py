#!/usr/bin/env python

#==============================================================================
#
# get_job_submitters.py - Output number of unique users submitting jobs.
#
# ARGS:
#  All: BillingDetails files
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
#===============================================================================

#=====
#
# IMPORTS
#
#=====
import argparse
from collections import defaultdict
import datetime
import os
import os.path
import xlrd
import sys

# Simulate an "include billing_common.py".
SCRIPT_DIR = os.path.dirname(os.path.realpath(__file__))
execfile(os.path.join(SCRIPT_DIR, "..", "billing_common.py"))

global sheet_get_named_column

#=====
#
# CONSTANTS
#
#=====

#=====
#
# FUNCTIONS
#
#=====

#=====
#
# SCRIPT BODY
#
#=====

parser = argparse.ArgumentParser()

parser.add_argument("billing_details_files", nargs="+")
parser.add_argument("-v", "--verbose", action="store_true",
                    default=False,
                    help='Get chatty [default = false]')
parser.add_argument("-d", "--debug", action="store_true",
                    default=False,
                    help='Get REAL chatty [default = false]')

args = parser.parse_args()

# Initialize data structures.
grand_total_user_set = set()

# Mapping from input file to tuple of unique users: (billable, nonbillable, failed, total)
files_to_user_tuple_map = dict()
# Mapping from input file to tuple of jobs: (billable, nonbillable, failed, total)
files_to_jobs_tuple_map = dict()

for billing_details_file in args.billing_details_files:

    # Reinitialize data structures for each file.
    total_user_set = set()
    billable_user_set = set()
    unbillable_user_set = set()
    failed_user_set = set()

    total_user_list = []
    billable_user_list = []
    unbillable_user_list = []
    failed_user_list = []

    # Open the BillingDetails workbook.
    #print "Reading BillingDetails workbook %s." % billing_details_file
    billing_details_wkbk = xlrd.open_workbook(billing_details_file, on_demand=True)

    # Get the users from the Computing sheet.
    computing_sheet = billing_details_wkbk.sheet_by_name('Computing')

    billable_user_list = sheet_get_named_column(computing_sheet, 'Username')

    billable_user_set.update(billable_user_list)

    # Get the users from the Unbillable Jobs sheet.
    # (try-except block guards against early BillingDetails files which didn't have this sheet)
    try:
        unbillable_sheet = billing_details_wkbk.sheet_by_name('Nonbillable Jobs')

        unbillable_user_list = sheet_get_named_column(unbillable_sheet, 'Username')

        unbillable_user_set.update(unbillable_user_list)
    except xlrd.biffh.XLRDError: pass

    # Get the users from the Failed Jobs sheet.
    # (try-except block guards against early BillingDetails files which didn't have this sheet)
    try:
        failed_sheet = billing_details_wkbk.sheet_by_name('Failed Jobs')

        failed_user_list = sheet_get_named_column(failed_sheet, 'Username')

        failed_user_set.update(failed_user_list)
    except xlrd.biffh.XLRDError: pass

    # Update the aggregate user set from the sheet user sets.
    total_user_set.update(billable_user_set)
    total_user_set.update(unbillable_user_set)
    total_user_set.update(failed_user_set)

    grand_total_user_set.update(total_user_set)

    total_user_list = billable_user_list + unbillable_user_list + failed_user_list

    files_to_user_tuple_map[billing_details_file] = (len(billable_user_set), len(unbillable_user_set), len(failed_user_set), len(total_user_set))
    files_to_jobs_tuple_map[billing_details_file] = (len(billable_user_list), len(unbillable_user_list), len(failed_user_list), len(total_user_list))

    billing_details_wkbk.release_resources()

# Print the unique user table.
print "Billable Users\tNonbillable Users\tFailed Users\tTotal Users"
for details_file in sorted(files_to_user_tuple_map.keys()):

    (billable_users, unbillable_users, failed_users, total_users) = files_to_user_tuple_map[details_file]
    print "%d\t%d\t%d\t%d\t%s" % (billable_users, unbillable_users, failed_users, total_users, details_file)

print

# Print the jobs table.
print "Billable Jobs\tNonbillable Jobs\tFailed Jobs\tTotal Jobs"
for details_file in sorted(files_to_user_tuple_map.keys()):

    (billable_jobs, unbillable_jobs, failed_jobs, total_jobs) = files_to_jobs_tuple_map[details_file]
    print "%d\t%d\t%d\t%d\t%s" % (billable_jobs, unbillable_jobs, failed_jobs, total_jobs, details_file)

print

# Print the total of all job submitting users.
print >> sys.stderr, len(grand_total_user_set),"\ttotal unique job submitting users"