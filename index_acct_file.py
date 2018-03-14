#!/usr/bin/env python

#===============================================================================
#
# NAMEOFSCRIPT.py - Description of script
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
import datetime
import os
import os.path
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

parser.add_argument("-a", "--accounting_file",
                    default=SGE_ACCOUNTING_FILE,
                    help='The SGE accounting file to snapshot [default = %s]' % SGE_ACCOUNTING_FILE)
parser.add_argument("--tuple", action="store_true",
                    default=False,
                    help='Output index table as a tuple of tuples [default = false]')
parser.add_argument("-v", "--verbose", action="store_true",
                    default=False,
                    help='Get real chatty [default = false]')

args = parser.parse_args()

year_month_to_filepos = dict()

#
# Open the current accounting file for input.
#
accounting_input_fp = open(args.accounting_file, "r", 0) # No buffering.

last_file_pos = 0
while True:

    line = accounting_input_fp.readline()
    if line == "": break
    if line[0] == "#": continue

    fields = line.split(':')
    submission_date = int(fields[8])
    end_date = int(fields[10])
    failed = int(fields[11])

    if failed in ACCOUNTING_FAILED_CODES:
        job_date = submission_date
    else:
        job_date = end_date

    job_date_date = datetime.date.fromtimestamp(job_date)
    year = int(job_date_date.strftime("%Y"))
    month = int(job_date_date.strftime("%m"))

    this_file_pos = accounting_input_fp.tell()

    year_month_tuple = (year, month)
    if year_month_tuple not in year_month_to_filepos:
        year_month_to_filepos[(year, month)] = last_file_pos

    last_file_pos = this_file_pos

#
# Print out the index table.
#
if args.tuple:
    print '(',
    for year_month_tuple in sorted(year_month_to_filepos.iterkeys()):
        (year, month) = year_month_tuple
        print "((%d, %d), %d)," % (year, month, year_month_to_filepos[year_month_tuple])
    print ')'

else:
    for year_month_tuple in sorted(year_month_to_filepos.iterkeys()):
        (year, month) = year_month_tuple
        print year, month, year_month_to_filepos[year_month_tuple]





