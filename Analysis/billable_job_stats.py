#!/usr/bin/env python

#===============================================================================
#
# billable_job_stats.py - Prints statistics about a user's job usage for a given month.
#
# ARGS:
#   None
#
# SWITCHES:
#   --user:            Username to print statistics for (default=current user).
#   --all_users:       Analyze jobs for all users (default=False).
#
#   --print_billable_jobs:     Print details of billable jobs to STDOUT (default=False).
#   --print_nonbillable_jobs:  Print details of nonbillable jobs to STDOUT (default=False).
#   --print_completed_jobs:    Print details of completed jobs to STDOUT (default=False).
#                                Note: completed jobs are billable jobs + nonbillable jobs.
#   --print_failed_jobs:       Print details of failed jobs to STDOUT (default=False).
#
#   --accounting_file: Location of accounting file (overrides BillingConfig.xlsx)
#   --billing_root:    Location of BillingRoot directory (overrides BillingConfig.xlsx)
#                      [default if no BillingRoot in BillingConfig.xlsx or switch: CWD]
#   --year:            Year of snapshot requested. [Default is this year]
#   --month:           Month of snapshot requested. [Default is this month]

#
# OUTPUT:
#    A table of statistics from the user's jobs for the given month.
#      Statistics include: CPU-hrs, job count, billable value.
#      Categories of jobs: billable, nonbillable, failed.
#    If any of the --print switches are present, lines about the particular
#      details of the job categories requested.
#
# ASSUMPTIONS:
#    None
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
from collections import defaultdict
import datetime
import argparse
import os
import os.path
import pwd
import sys


#=====
#
# CONSTANTS
#
#=====
NONBILLABLE_JOBS_EXIST = False   # If True, countenance the existence of nonbillable jobs.

SGE_ACCOUNTING_FILE = "/srv/gsfs0/admin_stuff/soge-8.1.8/scg4-feb2016/common/accounting"

# OGE accounting failed codes which invalidate the accounting entry.
# From http://docs.oracle.com/cd/E19080-01/n1.grid.eng6/817-6117/chp11-1/index.html
ACCOUNTING_FAILED_CODES = (1,3,4,5,6,7,8,9,10,11,26,27,28)

BILLING_RATE        = 0.08 # per CPU-hr
BILLING_FIRST_MONTH = 9    # Months before Sept 2013
BILLING_FIRST_YEAR  = 2013 #  were not billed.

USER = pwd.getpwuid(os.getuid()).pw_name

#=====
#
# FUNCTIONS
#
#=====
def from_ymd_date_to_timestamp(year, month, day):
    return int(calendar.timegm(datetime.date(year, month, day).timetuple()))

#=====
#
# SCRIPT BODY
#
#=====

parser = argparse.ArgumentParser()

parser.add_argument("-u","--user",
                    default=USER,
                    help="The user to compute statistics for [default = %s]" % USER)
parser.add_argument("-j","--job_tag",
                    default=None,
                    help="The job tag to compute statistics for [default = None]")
parser.add_argument("--all_users", action="store_true",
                    default=False,
                    help="Don't filter on users [default = False]")
parser.add_argument("--print_billable_jobs", action="store_true",
                    default=False,
                    help="Print details about billable jobs to STDOUT [default = False]")
if NONBILLABLE_JOBS_EXIST:
    parser.add_argument("--print_nonbillable_jobs", action="store_true",
                        default=False,
                        help="Print details about nonbillable jobs to STDOUT [default = False]")
parser.add_argument("--print_completed_jobs", action="store_true",
                    default=False,
                    help="Print details about completed jobs to STDOUT [default = False]")
parser.add_argument("--print_failed_jobs", action="store_true",
                    default=False,
                    help="Print details about failed jobs to STDOUT [default = False]")
parser.add_argument("-a", "--accounting_file",
                    default=SGE_ACCOUNTING_FILE,
                    help='The SGE accounting file to analyze [default = %s]' % SGE_ACCOUNTING_FILE)
parser.add_argument("-v", "--verbose", action="store_true",
                    default=False,
                    help='Get real chatty [default = false]')
parser.add_argument("-y","--year", type=int, choices=range(2012,2021),
                    default=None,
                    help="The year to be filtered out. [default = this year]")
parser.add_argument("-m", "--month", type=int, choices=range(1,13),
                    default=None,
                    help="The month to be filtered out. [default = this month]")

args = parser.parse_args()

#
# Sanity-check arguments.
#

if args.year is None:
    # No year given: use this year.
    year = datetime.date.today().year
else:
    year = args.year

if args.month is None:
    # No month given: use this month.
    month = datetime.date.today().month
else:
    month = args.month

# Calculate next month for range of this month.
if month != 12:
    next_month = month + 1
    next_month_year = year
else:
    next_month = 1
    next_month_year = year + 1

# Was this month billable?  (9/2013 or beyond)
is_billable_month = year > BILLING_FIRST_YEAR or (year == BILLING_FIRST_YEAR and month >= BILLING_FIRST_MONTH)

# The begin_ and end_month_timestamps are to be used as follows:
#   date is within the month if begin_month_timestamp <= date < end_month_timestamp
# Both values should be UTC.
begin_month_timestamp = from_ymd_date_to_timestamp(year, month, 1)
end_month_timestamp   = from_ymd_date_to_timestamp(next_month_year, next_month, 1)

# Get list of users from users arg.
if not args.all_users:
    user_list = args.user.split(',')
else:
    user_list = ["ALLUSERS"]

# Print first lines of upcoming table.
if args.all_users:
    table_header_str = "JOBS RUN BY ALL USERS"
else:
    table_header_str = "JOBS RUN BY USER %s" % (args.user)

if args.job_tag is not None:
    table_header_str += " WITH JOB TAG %s" % (args.job_tag)

print >> sys.stderr, "%s:" % table_header_str

if is_billable_month:
    print >> sys.stderr, "MONTH: %02d/%d\t\tCPU-hrs\t%7s\tCost" % (month, year, 'Jobs')
else:
    print >> sys.stderr, "MONTH: %02d/%d\t\tCPU-hrs\t%7s" % (month, year, 'Jobs')

#
# Read all the lines of the current accounting file.
#  Take statistics on all those lines
#  which have "end_times" in the given month.
#
with open(args.accounting_file, "r") as accounting_input_fp:

    this_month_user_jobs = defaultdict(list)
    this_month_failed_jobs = defaultdict(list)

    for line in accounting_input_fp:

        if line[0] == "#": continue

        fields = line.split(':')
        hostname = fields[1]
        owner = fields[3]
        job_name = fields[4]
        job_ID = fields[5]
        account = fields[6]
        submission_date = int(fields[8])
        end_date = int(fields[10])
        failed = int(fields[11])
        wallclock = int(fields[13])
        slots = int(fields[34])

        # Trim off trailing ".local" from hostname, if present.
        if hostname.endswith(".local"):
            hostname = hostname[:-6]

        # If this job failed, then use its submission_time as the job date.
        # else use the end_time as the job date.
        job_failed = failed in ACCOUNTING_FAILED_CODES
        if job_failed:
            job_date = submission_date  # No end_date for failed jobs.
        else:
            job_date = end_date

        # If the date of this job was within the month,
        #  save it for statistics.
        if begin_month_timestamp <= job_date < end_month_timestamp:

            job_date_string = datetime.datetime.utcfromtimestamp(job_date).strftime("%m/%d/%Y")

            # The job must be run by the requested user, if all_users not True.
            correct_user = args.all_users or owner in user_list
            # The job must match the given job tag, if any.
            correct_job_tag = args.job_tag is None or account == args.job_tag

            if correct_user and correct_job_tag:

                # Save the job details under "ALLUSERS" if args.all_users selected, else use the owner field.
                if args.all_users:
                    owner_or_allusers = "ALLUSERS"
                else:
                    owner_or_allusers = owner

                # Divide job details between successful and failed jobs.
                if not job_failed:
                    this_month_user_jobs[owner_or_allusers].append((hostname, owner, job_name, job_ID, job_date_string, account, slots, wallclock))
                else:
                    # One more failed job.
                    this_month_failed_jobs[owner_or_allusers].append((hostname, owner, job_name, job_ID, job_date_string, account, slots, wallclock, failed))


#
# Generate statistics from the runs.
#   If billable month, generate billable CPU-hr, cost, and count, nonbillable CPU-hr and count, and failed count.
#   Else nonbillable month, generate completed CPU-hr and count, and failed count.
#
if is_billable_month:

    # Analyze jobs found for user within the month/year specified.
    user_total_cpu_hrs = defaultdict(float)
    user_billable_cpu_hrs = defaultdict(float)
    user_nonbillable_cpu_hrs = defaultdict(float)

    this_month_billable_user_jobs = defaultdict(list)
    this_month_nonbillable_user_jobs = defaultdict(list)

    for user in user_list:
        for job_details in this_month_user_jobs[user]:

            (hostname, owner, job_name, job_ID, job_date, account, slots, wallclock) = job_details

            # Calculate this job's CPUslot-hrs.
            cpu_hrs = slots * wallclock / 3600.0

            # Count billable jobs: hostname does not start with 'scg3-0' or 'greenie'.
            if not (hostname.startswith('greenie') or hostname.startswith('scg3-0')):
                this_month_billable_user_jobs[user].append(job_details)
                user_billable_cpu_hrs[user] += cpu_hrs
            else:
                this_month_nonbillable_user_jobs[user].append(job_details)
                user_nonbillable_cpu_hrs[user] += cpu_hrs

            user_total_cpu_hrs[user] += cpu_hrs

        #
        # Compute stats on billable/nonbillable jobs, and print a small table with the results.
        #
        user_billable_job_count = len(this_month_billable_user_jobs[user])
        user_nonbillable_job_count = len(this_month_nonbillable_user_jobs[user])
        user_failed_job_count = len(this_month_failed_jobs[user])

        user_total_job_count = len(this_month_user_jobs[user]) + user_failed_job_count

        billable_cost = user_billable_cpu_hrs[user] * BILLING_RATE

        #
        # Print rest of output table
        #
        print >> sys.stderr, " %8s\tBilled\t%7.1f\t%7d\t$%0.02f" % (user, user_billable_cpu_hrs[user], user_billable_job_count, billable_cost)
        if NONBILLABLE_JOBS_EXIST:
            print >> sys.stderr, " %8s\tNonbill\t%7.1f\t%7d\t%7s" % (user, user_nonbillable_cpu_hrs[user], user_nonbillable_job_count, "--")
        # print >> sys.stderr, " %s\tFailed\t\t%7s\t%7d\t%7s" % (user, "N/A", user_failed_job_count[user], "--")
        if NONBILLABLE_JOBS_EXIST:
            print >> sys.stderr, "%8s\tTOTAL\t\t%7.1f\t%7d\t$%0.02f" % (user, user_total_cpu_hrs[user], user_total_job_count, billable_cost)

else:
    # Not a billable month: just return stats on job that ran vs jobs which failed.
    user_completed_cpu_hrs = defaultdict(float)

    for user in user_list:
        for job_details in this_month_user_jobs[user]:

            (hostname, owner, job_name, job_ID, job_date, account, slots, wallclock) = job_details

            # Calculate this job's CPUslot-hrs.
            cpu_hrs = slots * wallclock / 3600.0

        user_completed_cpu_hrs[user] += cpu_hrs

        user_completed_job_count = len(this_month_user_jobs[user])
        user_failed_job_count = len(this_month_failed_jobs[user])
        user_total_job_count = user_completed_job_count + user_failed_job_count

        user_total_cpu_hrs = user_completed_cpu_hrs[user]

        #
        # Print rest of output table
        #
        print >> sys.stderr, " Completed\t%7.1f\t%6d" % (user_completed_cpu_hrs, user_completed_job_count)
        print >> sys.stderr, " Failed\t\t%7s\t%6d" % ("N/A", user_failed_job_count)
        print >> sys.stderr, "TOTAL\t\t%7.1f\t%6d" % (user_total_cpu_hrs, user_total_job_count)

#
# Print user job details to stdout, if requested.
#
if is_billable_month:
    if args.print_billable_jobs or args.print_completed_jobs:
        for user in user_list:
            for job_details in this_month_billable_user_jobs[user]:
                print ':'.join(map(lambda s: str(s), job_details))
    if NONBILLABLE_JOBS_EXIST:
        if args.print_nonbillable_jobs or args.print_completed_jobs:
            for user in user_list:
                for job_details in this_month_nonbillable_user_jobs[user]:
                    print ':'.join(map(lambda s: str(s), job_details))
else:
    if args.print_completed_jobs:
        for user in user_list:
            for job_details in this_month_user_jobs[user]:
                print ':'.join(map(lambda s: str(s), job_details))
if args.print_failed_jobs:
    for user in user_list:
        for job_details in this_month_failed_jobs[user]:
            print ':'.join(map(lambda s: str(s), job_details))
