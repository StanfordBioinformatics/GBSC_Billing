#!/usr/bin/env python

#===============================================================================
#
# check_config.py - Confirm that the BillingConfiguration workbook makes sense.
#
# ARGS:
#   1st: The BillingConfig workbook to be checked.
#
# SWITCHES:
#   None
#
# OUTPUT:
#   None, if the file is OK.  O/W, messages regarding inconsistencies.
#
# ASSUMPTIONS:
#   Depends on xlrd module.
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
from collections import defaultdict
import argparse
import os
import os.path
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
global BILLING_CONFIG_SHEET_COLUMNS

#=====
#
# FUNCTIONS
#
#=====
# From billing_common.py
global sheet_get_named_column


def check_sheets(wkbk):

    sheets_are_OK = True

    all_sheets = wkbk.sheet_names()

    # Check each sheet.
    for sheet_name in BILLING_CONFIG_SHEET_COLUMNS:

        # Does the sheet exist?
        if sheet_name not in all_sheets:
            print "check_sheets:  Missing sheet %s" % sheet_name
            sheets_are_OK = True

        # Is the sheet loaded?
        if not wkbk.sheet_loaded(sheet_name):
            print >> sys.stderr, "check_sheets: Can't load %s sheet." % (sheet_name)
            sheets_are_OK = True

        # Look at the sheet itself.
        sheet = wkbk.sheet_by_name(sheet_name)

        # Does the sheet have the expected headings?
        expected_headers = BILLING_CONFIG_SHEET_COLUMNS[sheet_name]
        sheet_headers = sheet.row_values(0)
        for header in expected_headers:
            if header not in sheet_headers:
                print >> sys.stderr, "check_sheets: Sheet %s: Expected header %s missing." % (sheet_name, header)
                sheets_are_OK = False

    return sheets_are_OK

def check_pctages(sheet, col_name):

    pctages_are_OK = True

    objects = sheet_get_named_column(sheet, col_name)
    pctages = sheet_get_named_column(sheet, '%age')

    # Create a mapping from an object to the list of percentages associated with it.
    object_pctages = defaultdict(list)
    for (obj, pct) in zip(objects, pctages):
        object_pctages[obj].append(pct)

    # Add all the percentages for each object, and confirm that the sum is 100.
    for object in object_pctages:
        pctage_list = map(float, object_pctages[object])
        total_pctage = sum(pctage_list)
        if total_pctage != 1.0:
            print "check_pctages: %s %s percentages add up to %d%%, not 100%%" % (col_name, object, total_pctage*100.0)
            pctages_are_OK = False

    return pctages_are_OK


def check_folders_sheet(wkbk):

    folders_sheet = wkbk.sheet_by_name('Folders')

    # Check that, in the Folders sheet, the %ages associated with each folder
    #  add up to 100%.
    folder_pctages_are_OK = check_pctages(folders_sheet, 'Folder')

    return folder_pctages_are_OK


def check_jobtags_sheet(wkbk):

    jobtags_sheet = wkbk.sheet_by_name('JobTags')

    # Check that, in the JobTags sheet, the %ages associated with each job tag
    #  add up to 100%.
    jobtag_pctages_are_OK = check_pctages(jobtags_sheet, 'Job Tag')

    return jobtag_pctages_are_OK


def check_users_sheet(wkbk):

    users_sheet = wkbk.sheet_by_name('Users')
    pi_sheet    = wkbk.sheet_by_name('PIs')

    #
    # Check the PI Tags associated with each user.
    #

    # Mapping from username to list of [pi_tag, date_added, date_removed].
    username_to_pi_tag_dates = dict()

    for user_row_idx in range(1,users_sheet.nrows):
        (username, _, _, pi_tag, _, date_added, date_removed) = users_sheet.row_values(user_row_idx)

        pi_tag_date_list = username_to_pi_tag_dates.get(username)
        if pi_tag_date_list is None:
            username_to_pi_tag_dates[username] = [[pi_tag, date_added, date_removed]]
        else:
            username_to_pi_tag_dates[username].append([pi_tag, date_added, date_removed])

    # Is every PI tag in the list of PI tags in the PIs sheet?
    users_pi_tags_are_OK = True
    pi_tag_list = sheet_get_named_column(pi_sheet, 'PI Tag')

    for uname in username_to_pi_tag_dates:

        pi_tag_date_list = username_to_pi_tag_dates[uname]

        for (pi_tag, date_added, date_removed) in pi_tag_date_list:
            if pi_tag not in pi_tag_list:
                print >> sys.stderr, "check_users_sheet: User %s has non-existent PI Tag %s" % (uname, pi_tag)
                users_pi_tags_are_OK = False

    # Check that, in the Users sheet, the %ages associated with each user
    #  add up to 100%.
    users_pctages_are_OK = check_pctages(users_sheet, 'Username')

    return users_pi_tags_are_OK and users_pctages_are_OK

#=====
#
# SCRIPT BODY
#
#=====

parser = argparse.ArgumentParser()

parser.add_argument("billing_config_file",
                    help='The BillingConfig file')
parser.add_argument("-v", "--verbose", action="store_true",
                    default=False,
                    help='Get real chatty [default = false]')

args = parser.parse_args()

billing_config_wkbk = xlrd.open_workbook(args.billing_config_file)

# TODO: check that all sheets have columns of the same number of rows (except maybe Folders!By Quota? )
# TODO: Confirm that all folders exist.

# Check that all the proper sheets are in the spreadsheet.
sheets_are_OK = check_sheets(billing_config_wkbk)

if sheets_are_OK:
    # Check Folders sheet.
    folders_sheet_is_OK = check_folders_sheet(billing_config_wkbk)

    # Check JobTags sheet.
    job_tag_sheet_is_OK = check_jobtags_sheet(billing_config_wkbk)

    # Check Users sheet.
    users_sheet_is_OK   = check_users_sheet(billing_config_wkbk)
else:
    folders_sheet_is_OK = job_tag_sheet_is_OK = users_sheet_is_OK = False

if not (sheets_are_OK and
        folders_sheet_is_OK and
        job_tag_sheet_is_OK and
        users_sheet_is_OK):
    print "Billing Configuration spreadsheet has some problems."
    sys.exit(-1)
else:
    sys.exit(0)