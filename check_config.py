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
import grp
import pwd
import os
import os.path
import subprocess
import sys
import time

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
global EARLIEST_VALID_DATE_EXCELDATE
global PI_PROJECT_ROOT_DIR
global from_excel_date_to_date_string
global from_timestamp_to_excel_date

#=====
#
# FUNCTIONS
#
#=====
# From billing_common.py
global sheet_get_named_column

#
# CHECKS:
#  o that all sheets are in the BilingConfig workbook
#  o for each sheet: does it have the proper headings/columns?
#
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


def check_dates_pctages(sheet, col_name):

    dates_pctages_are_OK = True

    objects     = sheet_get_named_column(sheet, col_name)
    pi_tags     = sheet_get_named_column(sheet, 'PI Tag')
    dates_added = sheet_get_named_column(sheet, 'Date Added')
    dates_remvd = sheet_get_named_column(sheet, 'Date Removed')
    pctages     = sheet_get_named_column(sheet, '%age')

    # If there are no %ages, pretend every object was 100%.
    if pctages is None:
        pctages = [1.0] * len(objects)

    #
    # Create a mapping from an object to a list of [PI Tag, added, removed, %age].
    #
    object_entries = defaultdict(list)
    # Also save the dates encountered for that object.
    object_date_set = defaultdict(set)
    for (obj, pi_tag, date_added, date_removed, pctage) in zip(objects, pi_tags, dates_added, dates_remvd, pctages):

        # Check: is the removed date before or the same as the added date?
        if date_removed != '' and date_removed <= date_added:
            print "check_dates_pctages: %s %s date removed is before or equal to date added (%s < %s)" % (col_name, obj, from_excel_date_to_date_string(date_removed), from_excel_date_to_date_string(date_added))
            dates_pctages_are_OK = False

        # Check: is either date before the earliest possible date?
        if date_added < EARLIEST_VALID_DATE_EXCELDATE:
            print "check_dates_pctages: %s %s date added is before earliest possible date (%s < %s)" % (col_name, obj, from_excel_date_to_date_string(date_added), from_excel_date_to_date_string(EARLIEST_VALID_DATE_EXCELDATE))
            dates_pctages_are_OK = False
        if date_removed != '' and date_removed < EARLIEST_VALID_DATE_EXCELDATE:
            print "check_dates_pctages: %s %s date removed is before earliest possible date (%s < %s)" % (col_name, obj, from_excel_date_to_date_string(date_removed), from_excel_date_to_date_string(EARLIEST_VALID_DATE_EXCELDATE))
            dates_pctages_are_OK = False

        # Check: is the pctage less than 0?
        # Check: is the pctage greater than 100%?
        if pctage < 0.0:
            print "check_dates_pctages: %s %s %%age %.0f%% is less than 0%%" % (col_name, obj, float(pctage))
            dates_pctages_are_OK = False
        if pctage > 1.0:
            print "check_dates_pctages: %s %s %%age %.0f%% is greater than 100%%" % (col_name, obj, float(pctage))
            dates_pctages_are_OK = False

        # Save objects and their data.
        object_entries[obj].append((pi_tag, date_added, date_removed, pctage))

        # Save the dates.
        object_date_set[obj].add(date_added)
        if date_removed != '':
            object_date_set[obj].add(date_removed)

    # For all objects:
    #  For each of the dates important within their entries:
    #     Sum up the %ages for that date.
    for obj in object_entries:

        for (pi_tag, date_added, date_removed, pctage) in object_entries[obj]:

            total_pctage    = defaultdict(float)
            date_state_list = defaultdict(list)
            for date in object_date_set[obj]:

                # If this entry's date range includes this date:
                #   save a list of [PI_tag, %age] associated with that date.
                if date_added <= date < date_removed:
                    date_state_list[date].append((pi_tag, pctage))
                    total_pctage[date] += pctage

            for date in object_date_set[obj]:
                if total_pctage[date] > 1.0:
                    date_state_string = ""
                    for date_state in date_state_list[date]:
                        date_state_string += "(%s %d%%) " % (date_state[0], float(date_state[1]))
                    print "check_dates_pctages: %s %s: On %s, PI breakdown is >100%%:" % (col_name, obj, date), date_state_string

                    dates_pctages_are_OK = False

    return dates_pctages_are_OK


def check_pi_tags(sheet, col_name, pi_tag_list):

    pi_tags_are_OK = True

    objects = sheet_get_named_column(sheet, col_name)
    pi_tags = sheet_get_named_column(sheet, 'PI Tag')

    # Create a mapping from an object to the list of PI tags associated with it.
    object_pitags = defaultdict(set)
    for (obj, pi_tag) in zip(objects, pi_tags):
        object_pitags[obj].add(pi_tag)

    # Check all the PI Tags for the object, and whine if we don't know about all of them.
    for object in object_pitags:
        for pi_tag in object_pitags[object]:
            if pi_tag not in pi_tag_list:
                print "check_pi_tags: %s %s PI Tag %s not known." % (col_name, object, pi_tag)
                pi_tags_are_OK = False

    return pi_tags_are_OK


def check_usernames(sheet, col_name):

    usernames = sheet_get_named_column(sheet, col_name)
    fullnames = sheet_get_named_column(sheet, 'Full Name')

    usernames_all_OK = True

    # Create username -> fullname mapping.
    username_to_fullname = dict()
    for (username, fullname) in zip(usernames, fullnames):
        username_to_fullname[username] = fullname

    # Get list of currently active users for filtering.
    current_user_list = get_valid_objs_by_date(sheet, col_name, today)

    # Check all users for existence.
    for username in username_to_fullname:

        # Ignore invalid users.
        if username not in current_user_list:
            continue

        fullname = username_to_fullname[username]

        try:
            _ = pwd.getpwnam(username)
        except KeyError:
            print "check_username: User %s has unknown username '%s'" % (fullname, username)
            usernames_all_OK = False

    return usernames_all_OK


def check_groups(sheet, current_pi_tag_list):

    pi_tags = sheet_get_named_column(sheet, 'PI Tag')
    groups  = sheet_get_named_column(sheet, 'Group Name')

    groups_all_OK = True

    for (pi_tag, group) in zip(pi_tags, groups):

        # Ignore non-current PI tags.
        if pi_tag not in current_pi_tag_list:
            continue

        # "none" is a valid group name; check for it first.
        if group == "none":
            continue

        try:
            _ = grp.getgrnam(group)
        except KeyError:
            print "check_groups: PI Tag %s has invalid group name '%s'" % (pi_tag, group)
            groups_all_OK = False

    return groups_all_OK


def check_folders(current_folder_list):

    # Convert folder list to a set to eliminate duplicates.
    folder_set = set(current_folder_list)

    folders_all_OK = True

    for folder in folder_set:

        # Skip folders named "None".
        if folder == "None": continue

        # Split folder into machine:dir components.
        if folder.find(':') >= 0:
            (machine, dir) = folder.split(':')
        else:
            machine = None
            dir = folder

        # Check on this machine if folder is local;
        #  o/w, do a stat via ssh if the folder lives somewhere else.
        if machine is None:
            # Does the folder currently exist in the system?
            if not os.path.exists(folder):
                print "check_folders: Folder %s does not exist" % (folder)
                folders_all_OK = False
        else:
            stat_cmd = ['ssh', machine, '[ -e %s ]' % dir]
            returncode = subprocess.call(stat_cmd)

            if returncode != 0:
                print "check_folders: Folder %s does not exist" % (folder)
                folders_all_OK = False


    return folders_all_OK


def check_folder_methods(sheet):

    folder_list = sheet_get_named_column(sheet, 'Folder')
    method_list = sheet_get_named_column(sheet, 'Method')

    folder_method_hash = defaultdict(list)

    # Check: do all folders have a method, and vice versa?
    if len(folder_list) != len(method_list):
        print "check_folder_methods: Not 1-to-1 mapping from folders to methods"
        return False

    for (folder, method) in zip(folder_list, method_list):
        folder_method_hash[folder].append(method)

    folder_methods_all_OK = True

    # For any folder which has multiple methods, confirm that they are all the same.
    for folder in folder_method_hash:
        method_set = set(folder_method_hash[folder])
        if len(method_set) > 1:
            print "check_folder_methods: Folder %s is measured in multiple ways: %s" % ",".join(method_set)
            folder_methods_all_OK = False
        else:
            if folder_method_hash[folder][0] not in ['quota', 'usage', 'none']:
                print "check_folder_methods: Folder %s measurement method is not correct: %s" % folder_method_hash[folder][0]
                folder_methods_all_OK = False

    return folder_methods_all_OK


def get_pi_tag_list(sheet):
    return sheet_get_named_column(sheet, 'PI Tag')


def get_valid_objs_by_date(sheet, obj_name, valid_date):

    objects     = sheet_get_named_column(sheet, obj_name)
    dates_added = sheet_get_named_column(sheet, 'Date Added')
    dates_remvd = sheet_get_named_column(sheet, 'Date Removed')

    obj_set = set()
    for (obj, date_added, date_removed) in zip(objects, dates_added, dates_remvd):

        if date_added <= valid_date and (date_removed == '' or valid_date < date_removed):
            obj_set.add(obj)

    return list(obj_set)


def check_pis_sheet(wkbk, current_pi_tag_list):

    if args.verbose:
        print "CHECKING PI SHEET"

    pis_sheet = wkbk.sheet_by_name('PIs')
    folder_col_name = 'PI Folder'

    # CHECK: all the groups are valid.
    if args.verbose:
        print " Checking PI groups"
    groups_OK = check_groups(pis_sheet, current_pi_tag_list)

    #
    # TODO: Check: is the email syntax correct?
    #
    #if args.verbose:
    #    print " Checking PI email addresses"


    # Get list of current PI folders.
    current_folder_list = get_valid_objs_by_date(pis_sheet, folder_col_name, today)

    # CHECK: all the PI folders exist.
    if args.verbose:
        print " Checking that all folders exist"
    folders_OK = check_folders(current_folder_list)

    # CHECK: all the dates and %ages are valid.
    if args.verbose:
        print " Checking PI dates and percentages."
    dates_pctages_OK = check_dates_pctages(pis_sheet, 'PI Tag')

    #
    # CHECK: All iLab Service Request ID are numbers and unique.
    #
    iLab_service_req_IDs = sheet_get_named_column(pis_sheet,"iLab Service Request ID")

    if iLab_service_req_IDs is not None:

        if args.verbose:
            print " Checking iLab Service Request IDs."

        iLab_service_req_IDs_nonnumeric = []

        # Look in IDs for any which are not numbers.
        for id in iLab_service_req_IDs:
            id_str = str(id)
            if id_str != '' and not id_str.isdigit():
                iLab_service_req_IDs_nonnumeric.append(id_str)

        iLab_service_req_IDs_numbers_OK = len(iLab_service_req_IDs_nonnumeric) > 0

        if not iLab_service_req_IDs_numbers_OK:
            print "  check_iLab_service_req_IDs: The following IDs are not numbers:"
            print "   ",
            for id in iLab_service_req_IDs_nonnumeric:
                print id,
            print
    else:
        iLab_service_req_IDs_numbers_OK = True  # No IDs, no problem.

    return (groups_OK and folders_OK and dates_pctages_OK and
            iLab_service_req_IDs_numbers_OK)


def check_folders_sheet(wkbk, pi_tag_list):

    if args.verbose:
        print "CHECKING FOLDERS SHEET"

    folders_sheet = wkbk.sheet_by_name('Folders')
    folder_col_name = 'Folder'

    # TODO: Check: All expected headers are present.

    # Get list of current folders.
    current_folder_list = get_valid_objs_by_date(folders_sheet, folder_col_name, today)

    # Check: all the folders exist.
    if args.verbose:
        print " Checking that all folders exist"
    folders_OK = check_folders(current_folder_list)

    # Check: all the PI tags are valid.
    if args.verbose:
        print " Checking that all PI tags are valid"
    pi_tags_OK = check_pi_tags(folders_sheet, folder_col_name, pi_tag_list)

    # Check: all the Methods for a folder are the same and are valid.
    if args.verbose:
        print " Checking that all measurement methods are the same for each folder"
    methods_OK = check_folder_methods(folders_sheet)

    # Check: all the dates and %ages are valid.
    if args.verbose:
        print " Checking that all dates and percentages are valid"
    dates_pctages_OK = check_dates_pctages(folders_sheet, folder_col_name)

    return (folders_OK and pi_tags_OK and
            #pi_folders_OK and
            methods_OK and dates_pctages_OK)


def check_users_sheet(wkbk, pi_tag_list):

    if args.verbose:
        print "CHECKING USERS SHEET"

    users_sheet = wkbk.sheet_by_name('Users')

    # TODO: Check: All expected headers are present.

    username_col_name = 'Username'

    # Check: the username is a real UNIX username.
    if args.verbose:
        print " Checking that all usernames exist"
    usernames_OK = check_usernames(users_sheet, username_col_name)

    #
    # TODO: Check: is the email syntax correct?
    #
    #if args.verbose:
    #    print " Checking users' email addresses"

    # Check: all the PI tags are valid.
    if args.verbose:
        print " Checking PI tags for users"
    pi_tags_OK = check_pi_tags(users_sheet, username_col_name, pi_tag_list)

    # Check: all the dates and %ages are valid.
    if args.verbose:
        print " Checking that all dates and percentages are valid"
    dates_pctages_OK = check_dates_pctages(users_sheet, username_col_name)

    return (usernames_OK and pi_tags_OK and dates_pctages_OK)


def check_jobtags_sheet(wkbk, pi_tag_list):

    if args.verbose:
        print "CHECKING JOBTAGS SHEET"

    jobtags_sheet = wkbk.sheet_by_name('JobTags')

    # TODO: Check: All expected headers are present.

    jobtags_col_name = 'Job Tag'

    # Check: all the PI tags are valid.
    if args.verbose:
        print " Checking PI tags for job tags"
    pi_tags_OK = check_pi_tags(jobtags_sheet, jobtags_col_name, pi_tag_list)

    # Check: all the dates and %ages are valid.
    if args.verbose:
        print " Checking that all dates and percentages are valid"
    dates_pctages_OK = check_dates_pctages(jobtags_sheet, jobtags_col_name)

    return (pi_tags_OK and dates_pctages_OK)


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

today = from_timestamp_to_excel_date(time.time())

billing_config_wkbk = xlrd.open_workbook(args.billing_config_file)

# TODO: check that all sheets have columns of the same number of rows (except maybe Folders!By Quota? )

# Check that all the proper sheets are in the spreadsheet.
sheets_are_OK = check_sheets(billing_config_wkbk)

if sheets_are_OK:
    # Get list of pi_tags.
    pis_sheet = billing_config_wkbk.sheet_by_name('PIs')
    pi_tag_list = get_pi_tag_list(pis_sheet)

    # Get list of current pi_tags.
    current_pi_tag_list = get_valid_objs_by_date(pis_sheet, 'PI Tag', today)

    # Check PIs sheet.
    pis_sheet_is_OK     = check_pis_sheet(billing_config_wkbk, current_pi_tag_list)

    # Check Folders sheet.
    folders_sheet_is_OK = check_folders_sheet(billing_config_wkbk, pi_tag_list)

    # Check JobTags sheet.
    job_tag_sheet_is_OK = check_jobtags_sheet(billing_config_wkbk, pi_tag_list)

    # Check Users sheet.
    users_sheet_is_OK   = check_users_sheet(billing_config_wkbk, pi_tag_list)

#
# Output summary and set returncode.
#
if (sheets_are_OK and
    pis_sheet_is_OK and
    folders_sheet_is_OK and
    job_tag_sheet_is_OK and
    users_sheet_is_OK):

    print
    print "+++ BILLING CONFIGURATION CONFIRMED +++"
    sys.exit(0)
else:
    print
    print "*** PROBLEMS WITH BILLING CONFIGURATION ***"
    if not sheets_are_OK:
        print "  Sheets are missing."
    if not pis_sheet_is_OK:
        print "  Problems with PI sheet"
    if not folders_sheet_is_OK:
        print "  Problems with Folders sheet"
    if not job_tag_sheet_is_OK:
        print "  Problems with Job Tag sheet"
    if not users_sheet_is_OK:
        print "  Problems with Users sheet"

    sys.exit(-1)