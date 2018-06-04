#!/usr/bin/env python

# ===============================================================================
#
# analyze_job_count.py -
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
# ==============================================================================

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
import xlsxwriter
import sys

# Simulate an "include billing_common.py".
SCRIPT_DIR = os.path.dirname(os.path.realpath(__file__))
execfile(os.path.join(SCRIPT_DIR, "billing_common.py"))

global from_excel_date_to_date_string

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
def get_total_slots_wallclock(sheet):

    if sheet is None:
        return (0,0)

    total_slots = 0
    total_wallclk = 0
    for row in range(1, sheet.nrows):
        row_vals = sheet.row_values(row)
        if len(row_vals) == 9:
            (date, timestamp, username, job_name, account, node, cores, wallclk_secs, job_id) = row_vals
        elif len(row_vals) == 10:
            (date, timestamp, username, job_name, account, node, cores, wallclk_secs, job_id, _) = row_vals
        total_slots += cores
        total_wallclk += wallclk_secs

    return (total_slots, total_wallclk)


def read_billingdetails_file(bd_wkbk):

    computing_sheet = bd_wkbk.sheet_by_name("Computing")

    billable_jobs = computing_sheet.nrows - 1

    if "Nonbillable Jobs" in bd_wkbk.sheet_names():
        nonbillable_sheet = bd_wkbk.sheet_by_name("Nonbillable Jobs")
        nonbillable_jobs = nonbillable_sheet.nrows - 1
    else:
        nonbillable_sheet = None
        nonbillable_jobs = 0

    if "Failed Jobs" in bd_wkbk.sheet_names():
        failed_sheet = bd_wkbk.sheet_by_name("Failed Jobs")
        failed_jobs = failed_sheet.nrows - 1
    else:
        failed_sheet = None
        failed_jobs = 0

    total_jobs = billable_jobs + nonbillable_jobs + failed_jobs

    # Get slot counts and wallclocks
    (billable_slots, billable_wallclock) = get_total_slots_wallclock(computing_sheet)
    (nonbillable_slots, nonbillable_wallclock) = get_total_slots_wallclock(nonbillable_sheet)
    (failed_slots, failed_wallclock) = get_total_slots_wallclock(failed_sheet)

    total_slots = billable_slots + nonbillable_slots + failed_slots
    total_wallclock = billable_wallclock + nonbillable_wallclock + failed_wallclock

    print "%d\t%d\t%d\t%d\t%d\t%d\t%d\t%d\t%d\t%d\t%d\t%d" % (billable_jobs, nonbillable_jobs, failed_jobs, total_jobs,
                                                              billable_slots, nonbillable_slots, failed_slots, total_slots,
                                                              billable_wallclock, nonbillable_wallclock, failed_wallclock, total_wallclock)



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

print "Filename\tBillable Jobs\tNon-billable Jobs\tFailed Jobs\tTotal Jobs\tBillable Slots\tNon-billable Slots\tFailed Slots\tTotal Slots\tBillable Secs\tNon-billable Secs\tFailed Secs\tTotal Secs"

for billing_details_file in sorted(args.billing_details_files):

    if billing_details_file.endswith('.xlsx'):
        # Open the BillingDetails workbook.
        #print >> sys.stderr, "Reading BillingDetails workbook %s." % billing_aggregate_file

        print billing_details_file + "\t",

        billing_details_wkbk = xlrd.open_workbook(billing_details_file, on_demand=True)

        read_billingdetails_file(billing_details_wkbk)

        billing_details_wkbk.release_resources()

