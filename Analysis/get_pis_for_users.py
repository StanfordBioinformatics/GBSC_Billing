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
import fileinput
import os.path
import sys
import xlrd

# Simulate an "include billing_common.py".
SCRIPT_DIR = os.path.dirname(os.path.realpath(__file__))
execfile(os.path.join(SCRIPT_DIR, "..", "billing_common.py"))

# From billing_common.py
global sheet_get_named_column
global from_timestamp_to_excel_date

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
# Creates all the data structures used to write the BillingNotification workbook.
# The overall goal is to mimic the tables of the notification sheets so that
# to build the table, all that is needed is to print out one of these data structures.
def build_global_data(wkbk):

    users_sheet    = wkbk.sheet_by_name("Users")

    #
    # Create mapping from usernames to a list of pi_tag/dates.
    #
    global username_to_pi_tag_dates

    usernames  = sheet_get_named_column(users_sheet, "Username")
    pi_tags = sheet_get_named_column(users_sheet, "PI Tag")
    dates_added = sheet_get_named_column(users_sheet, "Date Added")
    dates_removed = sheet_get_named_column(users_sheet, "Date Removed")
    pctages = sheet_get_named_column(users_sheet, "%age")

    username_rows = zip(usernames, pi_tags, dates_added, dates_removed, pctages)

    for (username, pi_tag, date_added, date_removed, pctage) in username_rows:
        username_to_pi_tag_dates[username].append([pi_tag, date_added, date_removed, pctage])


# This function scans the username_to_pi_tag_dates dict to create a list of [pi_tag, %age]s
# for the PIs that the given user was working for on the given date.
def get_pi_tags_for_username_by_date(username, date_timestamp):

    # Add PI Tag to the list if the given date is after date_added, but before date_removed.

    pi_tag_list = []

    pi_tag_dates = username_to_pi_tag_dates.get(username)
    if pi_tag_dates is not None:

        try:
            date_excel = from_timestamp_to_excel_date(date_timestamp)
        except Exception as e:
            print date_timestamp
            raise e

        for (pi_tag, date_added, date_removed, pctage) in pi_tag_dates:
            if date_added <= date_excel < date_removed:
                pi_tag_list.append([pi_tag, pctage])

    return pi_tag_list

#=====
#
# SCRIPT BODY
#
#=====

parser = argparse.ArgumentParser()

parser.add_argument("-c","--billing_config_file",
                    default=None,
                    help="The BillingConfig file")
parser.add_argument("-v", "--verbose", action="store_true",
                    default=False,
                    help='Get chatty [default = false]')
parser.add_argument("-d", "--debug", action="store_true",
                    default=False,
                    help='Get REAL chatty [default = false]')

args = parser.parse_args()

if args.billing_config_file is None:
    print >> sys.stderr, "I need a BillingConfig file...exiting."
    sys.exit(-1)

# Open the BillingConfig workbook.
billing_config_wkbk = xlrd.open_workbook(args.billing_config_file, on_demand=True)
#
# Build data structures.
#
build_global_data(billing_config_wkbk)

# Find the Users sheet.
users_sheet = billing_config_wkbk.sheet_by_name("Users")

# Read a subtable from the Users sheet with just Users and Date Removed.
users = sheet_get_named_column(users_sheet, "Username")

for line in fileinput.input():
    line.rstrip()  # Remove trailing \n

    ## NOT FINISHED YET - I NEED DATES FOR THE USERS TO SEE WHO THEIR PIs WERE ON THOSE DATES

