#!/usr/bin/env python

#===============================================================================
#
# snapshot_accounting.py - Copies the given month/year's SGE accounting data
#                           into a separate file.
#
# ARGS:
#   1st: BillingConfig.xlsx file (for Config sheet: location of accounting file)
#        [optional if --accounting_file given]
#   2nd: month as number
#        [optional: if not present, last month will be used.]
#   3rd: year [optional: if not present, current year will be used.]
#
# SWITCHES:
#   --accounting_file: Location of accounting file (overrides BillingConfig.xlsx)
#   --billing_root:    Location of BillingRoot directory (overrides BillingConfig.xlsx)
#                      [default if no BillingRoot in BillingConfig.xlsx or switch: CWD]
#
# OUTPUT:
#    An accounting file with only entries with submission_dates within the given
#     month.  This file, named SGEAccounting.<YEAR>-<MONTH>.txt, will be placed in
#     <BillingRoot>/<YEAR>/<MONTH>/ if BillingRoot is given or in the current
#     working directory if not.
#
# ASSUMPTIONS:
#    Directory hierarchy <BillingRoot>/<YEAR>/<MONTH>/ already exists.
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
import time
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
SGE_ACCOUNTING_FILE = "/srv/gs1/software/oge2011.11p1/scg3-oge-new/common/accounting"

# OGE accounting failed codes which invalidate the accounting entry.
# From http://docs.oracle.com/cd/E19080-01/n1.grid.eng6/817-6117/chp11-1/index.html
ACCOUNTING_FAILED_CODES = (1,3,4,5,6,7,8,9,10,11,26,27,28)

BILLING_RATE        = 0.10 # per CPU-hr
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
parser.add_argument("--all_users", action="store_true",
                    default=False,
                    help="Don't filter on users [default = False]")
parser.add_argument("--print_billable_jobs", action="store_true",
                    default=False,
                    help="Print details about billable jobs to STDOUT [default = False]")
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
                    help='The SGE accounting file to snapshot [default = %s]' % SGE_ACCOUNTING_FILE)
parser.add_argument("-v", "--verbose", action="store_true",
                    default=False,
                    help='Get real chatty [default = false]')
parser.add_argument("-y","--year", type=int, choices=range(2012,2021),
                    default=None,
                    help="The year to be filtered out. [default = this year]")
parser.add_argument("-m", "--month", type=int, choices=range(1,13),
                    default=None,
                    help="The month to be filtered out. [default = last month]")

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

# Print first lines of upcoming table.
if args.all_users:
    print >> sys.stderr, "JOBS RUN BY ALL USERS:"
else:
    print >> sys.stderr, "JOBS RUN BY USER %s:" % (args.user)
if is_billable_month:
    print >> sys.stderr, "MONTH: %02d/%d\tCPU-hrs\tJobs\tCost" % (month, year)
else:
    print >> sys.stderr, "MONTH: %02d/%d\tCPU-hrs\tJobs" % (month, year)

#
# Read all the lines of the current accounting file.
#  Take statistics on all those lines
#  which have "end_times" in the given month.
#
with open(args.accounting_file, "r") as accounting_input_fp:

    this_month_user_jobs = []
    this_month_failed_jobs = []

    for line in accounting_input_fp:

        if line[0] == "#": continue

        fields = line.split(':')
        hostname = fields[1]
        owner = fields[3]
        job_name = fields[4]
        job_ID = fields[5]
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

            # The job must also be run by the requested user, if all_users not True.
            if args.all_users or owner == args.user:
                if not job_failed:
                    this_month_user_jobs.append((hostname, owner, job_name, job_ID, job_date_string, slots, wallclock))
                else:
                    # One more failed job.
                    this_month_failed_jobs.append((hostname, owner, job_name, job_ID, job_date_string, slots, wallclock, failed))


#
# Generate statistics from the runs.
#   If billable month, generate billable CPU-hr, cost, and count, nonbillable CPU-hr and count, and failed count.
#   Else nonbillable month, generate completed CPU-hr and count, and failed count.
#
if is_billable_month:

    # Analyze jobs found for user within the month/year specified.
    user_total_cpu_hrs = 0
    user_billable_cpu_hrs = 0
    user_nonbillable_cpu_hrs = 0

    this_month_billable_user_jobs = []
    this_month_nonbillable_user_jobs = []

    for job_details in this_month_user_jobs:

        (hostname, owner, job_name, job_ID, end_date, slots, wallclock) = job_details

         # Calculate this job's CPUslot-hrs.
        cpu_hrs = slots * wallclock / 3600.0

        # Count billable jobs: hostname starts with 'scg1'.
        if hostname.startswith('scg1'):
            this_month_billable_user_jobs.append(job_details)
            user_billable_cpu_hrs += cpu_hrs
        else:
            this_month_nonbillable_user_jobs.append(job_details)
            user_nonbillable_cpu_hrs += cpu_hrs

        user_total_cpu_hrs += cpu_hrs

    #
    # Compute stats on billable/nonbillable jobs, and print a small table with the results.
    #
    user_billable_job_count = len(this_month_billable_user_jobs)
    user_nonbillable_job_count = len(this_month_nonbillable_user_jobs)
    user_failed_job_count = len(this_month_failed_jobs)

    user_total_job_count = len(this_month_user_jobs) + user_failed_job_count

    billable_cost = user_billable_cpu_hrs * BILLING_RATE

    #
    # Print rest of output table
    #
    print >> sys.stderr, " Billable\t%7.1f\t%6d\t$%0.02f" % (user_billable_cpu_hrs, user_billable_job_count, billable_cost)
    print >> sys.stderr, " Nonbillable\t%7.1f\t%6d\t%6s" % (user_nonbillable_cpu_hrs, user_nonbillable_job_count, "--")
    print >> sys.stderr, " Failed\t\t%7s\t%6d\t%6s" % ("N/A", user_failed_job_count, "--")
    print >> sys.stderr, "TOTAL\t\t%7.1f\t%6d\t$%0.02f" % (user_total_cpu_hrs, user_total_job_count, billable_cost)

else:
    # Not a billable month: just return stats on job that ran vs jobs which failed.
    user_total_cpu_hrs = 0

    for job_details in this_month_user_jobs:

        (hostname, owner, job_name, job_ID, end_date, slots, wallclock) = job_details

        # Calculate this job's CPUslot-hrs.
        cpu_hrs = slots * wallclock / 3600.0

        user_total_cpu_hrs += cpu_hrs

    user_completed_job_count = len(this_month_user_jobs)
    user_failed_job_count = len(this_month_failed_jobs)
    user_total_job_count = user_completed_job_count + user_failed_job_count

    #
    # Print rest of output table
    #
    print >> sys.stderr, " Completed\t%7.1f\t%6d" % (user_total_cpu_hrs, user_total_job_count)
    print >> sys.stderr, " Failed\t\t%7s\t%6d" % ("N/A", user_failed_job_count)
    print >> sys.stderr, "TOTAL\t\t%7.1f\t%6d" % (user_total_cpu_hrs, user_total_job_count)

#
# Print user job details to stdout, if requested.
#
if is_billable_month:
    if args.print_billable_jobs or args.print_completed_jobs:
        for job_details in this_month_billable_user_jobs:
            print ':'.join(map(lambda s: str(s), job_details))
    if args.print_nonbillable_jobs or args.print_completed_jobs:
        for job_details in this_month_nonbillable_user_jobs:
            print ':'.join(map(lambda s: str(s), job_details))
else:
    if args.print_completed_jobs:
        for job_details in this_month_user_jobs:
            print ':'.join(map(lambda s: str(s), job_details))
if args.print_failed_jobs:
    for job_details in this_month_failed_jobs:
        print ':'.join(map(lambda s: str(s), job_details))
