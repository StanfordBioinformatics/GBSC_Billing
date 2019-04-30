#!/usr/bin/env python

#===============================================================================
#
# get_uv300_hours.py - Determines how many hours within a month some UV-300 job was running.
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
import argparse
import os
import os.path
import sys
import datetime
import calendar
import csv
import time

import job_accounting_file

# Simulate an "include billing_common.py".
SCRIPT_DIR = os.path.dirname(os.path.realpath(__file__))
execfile(os.path.join(SCRIPT_DIR, "..", "billing_common.py"))

#=====
#
# CONSTANTS
#
#=====
# from billing_common.py
global SLURMACCOUNTING_PREFIX

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

parser.add_argument("-s", "--slurm_accounting_file",
                    default=None,
                    help='The Slurm accounting file to read [default = None]')
parser.add_argument("-r", "--billing_root",
                    default=None,
                    help='The Billing Root directory [default = None]')
parser.add_argument("-y","--year", type=int, choices=range(2013,2031),
                    default=None,
                    help="The year to be filtered out. [default = this year]")
parser.add_argument("-m", "--month", type=int, choices=range(1,13),
                    default=None,
                    help="The month to be filtered out. [default = last month]")

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

first_day_of_this_month = datetime.datetime(year,month,1)

billing_root = args.billing_root
# If we still don't have a billing root dir, use the current directory.
if billing_root is None:
    billing_root = os.getcwd()

# Get absolute path for billing_root directory.
billing_root = os.path.abspath(billing_root)

# Within BillingRoot, create YEAR/MONTH dirs if necessary.
year_month_dir = os.path.join(billing_root, str(year), "%02d" % month)
if not os.path.exists(year_month_dir):
    os.makedirs(year_month_dir)

# Use switch arg for accounting_file if present, else use file in BillingRoot.
if args.slurm_accounting_file is not None:
    accounting_file = args.slurm_accounting_file
else:
    accounting_filename = "%s.%d-%02d.txt" % (SLURMACCOUNTING_PREFIX, year, month)
    accounting_file = os.path.join(year_month_dir, accounting_filename)

# Get absolute path for accounting_file.
accounting_file = os.path.abspath(accounting_file)

print "  Accounting File: %s" % accounting_file

#
# Create data structure for all hours within month.
#
(_, days_in_month) = calendar.monthrange(year, month)
jobs_per_hour = [[0 for x in range(24)] for y in range(days_in_month)]
total_hours_in_month = days_in_month * 24

total_jobs = 0

#
# Read the Slurm accounting file.
#

# Read in the header line from the Slurm file to use for the DictReader
slurm_fp = open(accounting_file, "r")
header = slurm_fp.readline()
fieldnames = header.split('|')
reader = csv.DictReader(slurm_fp, fieldnames=fieldnames, dialect="slurm")

one_hour = datetime.timedelta(hours=1)

#   For lines which include "uv300" in the NodeList:
#     For each hour from 'start' to 'end':
#       Increment jobs_per_hour for that [day,hour].
#
for slurm_job in reader:

    if 'sgiuv300' in slurm_job['NodeList']:
        total_jobs += 1

        start_date = datetime.datetime.strptime(slurm_job['Start'], "%Y-%m-%dT%H:%M:%S")
        end_date   = datetime.datetime.strptime(slurm_job['End'], "%Y-%m-%dT%H:%M:%S")

        # Account for jobs which start in previous months.
        if start_date < first_day_of_this_month:
            start_date = first_day_of_this_month  # Only examining hours in this month.

        curr_date  = start_date
        while curr_date < end_date:
            curr_hour = curr_date.hour   # already 0-23
            curr_day  = curr_date.day-1  # need zero-indexed

            jobs_per_hour[curr_day][curr_hour] += 1

            curr_date += one_hour

    if total_jobs % 1000 == 0:
        sys.stderr.write('.')
        sys.stderr.flush()

print >> sys.stderr

print "Total hours per month: %d" % (days_in_month * 24)

total_hours_with_jobs = sum([1 for day in jobs_per_hour for hour in day if hour > 0])

print "Total hours with UV300 job: %d" % total_hours_with_jobs
print "%%age of hours with UV300 job: %3.1f%%" % (total_hours_with_jobs * 100.0 / total_hours_in_month)

print
print "Total UV300 jobs: %d" % total_jobs


