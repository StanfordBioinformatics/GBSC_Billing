#!/usr/bin/env python

# ===============================================================================
#
# get_user_list.py - Outputs a list of unique users from a BillingConfig file.
#
# ARGS:
#  1st: BillingConfig file
#
# SWITCHES:
#
# OUTPUT:
#
# ASSUMPTIONS:
#  All dates in the Date Removed columns are in the past -- all users with
#  Date Removed dates will be ignored.
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
import os
import os.path
import sys
import xlrd

# Simulate an "include billing_common.py".
SCRIPT_DIR = os.path.dirname(os.path.realpath(__file__))
execfile(os.path.join(SCRIPT_DIR, "billing_common.py"))

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

parser.add_argument("billing_config_file")
parser.add_argument("-v", "--verbose", action="store_true",
                    default=False,
                    help='Get chatty [default = false]')
parser.add_argument("-d", "--debug", action="store_true",
                    default=False,
                    help='Get REAL chatty [default = false]')

args = parser.parse_args()

# Open the BillingConfig workbook.
billing_details_wkbk = xlrd.open_workbook(args.billing_config_file, on_demand=True)

# Find the Users sheet.
users_sheet = billing_details_wkbk.sheet_by_name("Users")

# Read a subtable from the Users sheet with just Users and Date Removed.
users = sheet_get_named_column(users_sheet, "Username")
dates_removed = sheet_get_named_column(users_sheet, "Date Removed")

# Store the set of unique users in this set.
user_set = set()

for (user, date_removed) in zip(users, dates_removed):
    if date_removed == '':
        user_set.add(user)

# Print out the user list.
for user in sorted(user_set):
    print user

print >> sys.stderr, len(user_set), "\tusers in", args.billing_config_file
