#!/usr/bin/env python

#===============================================================================
#
# get_uv300_table.py - Produces table of UV-300 jobs.
#
# ARGS:
#    1st: BillingConfig file
#    2nd: SlurmAccounting file
# ...nth: SlurmAccounting file
#
# SWITCHES:
#
# OUTPUT:
#    To stdout: CSV table with columns of:
#       Start Date, JobID, User, Wallclock, Account, PI list

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
import os.path
import sys
import datetime

import xlrd

from job_accounting_file import JobAccountingFile

# Simulate an "include billing_common.py".
SCRIPT_DIR = os.path.dirname(os.path.realpath(__file__))
exec(compile(open(os.path.join(SCRIPT_DIR, "..", "billing_common.py"), "rb").read(), os.path.join(SCRIPT_DIR, "..", "billing_common.py"), 'exec'))

#=====
#
# CONSTANTS
#
#=====
# from billing_common.py
global from_timestamp_to_excel_date
global sheet_get_named_column

#=====
#
# GLOBALS
#
#=====
# Mapping from usernames to list of [date, pi_tag].
username_to_pi_tag_dates = defaultdict(list)

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

    username_rows = list(zip(usernames, pi_tags, dates_added, dates_removed, pctages))

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
            print(date_timestamp)
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

parser.add_argument("billing_config_file",
                    help='The BillingConfig file')
parser.add_argument("slurm_accounting_files", nargs='+',
                    help='Slurm Accounting files')
parser.add_argument("-v", "--verbose", action="store_true",
                    default=False,
                    help='Get chatty [default = false]')
parser.add_argument("-d", "--debug", action="store_true",
                    default=False,
                    help='Get REAL chatty [default = false]')

args = parser.parse_args()

#
# Process arguments.
#

###
#
# Read the BillingConfig workbook and build input data structures.
#
###

billing_config_wkbk = xlrd.open_workbook(args.billing_config_file)
#
# Build data structures.
#
build_global_data(billing_config_wkbk)

# print header for table
print("start_date,jobID,user,wallclock,cpus,account,pilist")

total_jobs = 0
total_uv300_jobs = 0
for accounting_file in args.slurm_accounting_files:

    # Get absolute path for accounting_file.
    accounting_file = os.path.abspath(accounting_file)

    print("  Accounting File: %s" % accounting_file, file=sys.stderr)

    total_uv300_jobs_in_file = 0

    #
    # Read the Slurm accounting file.
    #

    # Read in the header line from the Slurm file to use for the DictReader
    slurm_file = JobAccountingFile(accounting_file)

    #   For lines which include "uv300" in the NodeList:
    #
    for slurm_rec in slurm_file:
        total_jobs += 1

        if 'sgiuv300' in slurm_rec.node_list:
            total_uv300_jobs_in_file += 1

            start_date = slurm_rec.start_time
            jobID = slurm_rec.job_id
            user = slurm_rec.owner
            wallclock = slurm_rec.wallclock
            cpus = slurm_rec.cpus
            account = slurm_rec.account

            pi_pct_list = get_pi_tags_for_username_by_date(user, start_date)
            pi_list = list(set([x[0] for x in pi_pct_list]))
            pis = "+".join(pi_list)

            print(",".join(map(str,[start_date,jobID,user,wallclock,cpus,account,pis])))

            if total_jobs % 1000 == 0:
                sys.stderr.write('.')
                sys.stderr.flush()

    print(file=sys.stderr)
    print("  Total UV300 jobs for %s: %d" % (accounting_file, total_uv300_jobs_in_file), file=sys.stderr)

    total_uv300_jobs += total_uv300_jobs_in_file

print(file=sys.stderr)
print("Total UV300 jobs in all files: %d" % total_uv300_jobs, file=sys.stderr)
print("Total jobs in all files: %d" % total_jobs, file=sys.stderr)
