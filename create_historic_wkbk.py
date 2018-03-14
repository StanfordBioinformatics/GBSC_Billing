#!/usr/bin/env python

# ===============================================================================
#
# create_historic_wkbk.py - Compile historic workbook from BillingAggregate files.
#
# ARGS:
#  All: BillingAggregate files
#
# SWITCHES:
#
# OUTPUT:
#   An Excel workbook with sheets for Compute, Storage, Cloud, Consulting, and
#     Total charges.
#     Each sheet would be PI_TAGs vs months.
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
import re
import xlrd
import xlsxwriter
import sys

# Simulate an "include billing_common.py".
SCRIPT_DIR = os.path.dirname(os.path.realpath(__file__))
execfile(os.path.join(SCRIPT_DIR, "billing_common.py"))

# From billing_common.py
global sheet_get_named_column
global from_excel_date_to_date_string

#=====
#
# CONSTANTS
#
#=====
PI_TAG_COLUMN_NAME     = "PI Tag"
STORAGE_COLUMN_NAME    = "Storage"
COMPUTE_COLUMN_NAME    = "Computing"
CLOUD_COLUMN_NAME      = "Cloud"
CONSULTING_COLUMN_NAME = "Consulting"
TOTAL_CHGS_COLUMN_NAME = "Total Charges"

#=====
#
# FUNCTIONS
#
#=====
def read_billingaggregate_file(ba_wkbk, monthbilled):

    global pi_tag_set

    sheet = ba_wkbk.sheet_by_index(0)

    pi_tag_list = sheet_get_named_column(sheet, PI_TAG_COLUMN_NAME)

    # Every sheet has storage, compute, and total charges.
    storage_chgs = sheet_get_named_column(sheet, STORAGE_COLUMN_NAME)
    compute_chgs = sheet_get_named_column(sheet, COMPUTE_COLUMN_NAME)
    total_chgs   = sheet_get_named_column(sheet, TOTAL_CHGS_COLUMN_NAME)

    # Cloud and consulting charges may not be present.
    cloud_chgs   = sheet_get_named_column(sheet, CLOUD_COLUMN_NAME)
    if cloud_chgs is None:
        cloud_chgs = ["-"] * len(storage_chgs)
    consulting_chgs = sheet_get_named_column(sheet, CONSULTING_COLUMN_NAME)
    if consulting_chgs is None:
        consulting_chgs = ["-"] * len(storage_chgs)

    for (pi_tag, storage, compute, cloud, consulting, total) in zip(pi_tag_list, storage_chgs, compute_chgs, cloud_chgs, consulting_chgs, total_chgs):

        if pi_tag == '' or pi_tag == "TOTAL": break

        pi_tag_month_to_costs_dict[(pi_tag, monthbilled)] = (storage, compute, cloud, consulting, total)
        pi_tag_set.add(pi_tag)


def write_history_wkbk(out_wkbk, new_sheet_name, months_billed, tuple_index):

    # Create the new sheet to output the data in.
    sheet = out_wkbk.add_worksheet(new_sheet_name)

    pi_tags_sorted = sorted(pi_tag_set)
    months_sorted = sorted(months_billed)

    # Write the column headings: PI Tag and month dates.
    sheet.write_string(0, 0, PI_TAG_COLUMN_NAME, heading_format)
    sheet.write_row(0, 1, months_sorted, heading_format)

    # Write the first column of PI tags.
    sheet.write_column(1, 0, pi_tags_sorted)

    row = 1
    col = 0
    for pi_tag in pi_tags_sorted:

        # Write the PI tag in the first column.
        sheet.write_string(row, col, pi_tag)
        col += 1

        # Write the appropriate costs for the month in the subsequent columns.
        for month in months_sorted:
            cost_tuple = pi_tag_month_to_costs_dict.get((pi_tag, month))
            if cost_tuple is not None:
                if type(cost_tuple[tuple_index]) is not str:
                    sheet.write_number(row, col, cost_tuple[tuple_index], money_format)
                else:
                    sheet.write_string(row, col, cost_tuple[tuple_index])
            else:
                sheet.write_string(row, col, "n/a")
            col += 1

        row += 1
        col = 0


#=====
#
# SCRIPT BODY
#
#=====

parser = argparse.ArgumentParser()

parser.add_argument("billing_aggregate_files", nargs="+")
parser.add_argument("-o","--output",
                    default="GBSCBilling_History.xlsx",
                    help='Filename of resulting .xlsx file.')
parser.add_argument("-v", "--verbose", action="store_true",
                    default=False,
                    help='Get chatty [default = false]')
parser.add_argument("-d", "--debug", action="store_true",
                    default=False,
                    help='Get REAL chatty [default = false]')

args = parser.parse_args()

# Initialize data structures.

# Mapping from (pi_tag, month_date) to tuple of (storage_costs, compute_costs, cloud_costs, consulting_costs, total_costs)
pi_tag_month_to_costs_dict = dict()

# A set of all the pi_tags encountered.
pi_tag_set = set()

# A list of all the months in BillingAggregate files.
months_billed = []

for billing_aggregate_file in args.billing_aggregate_files:

    if billing_aggregate_file.endswith('.xlsx'):

        # Parse the name of the file to come up with the month and year.
        # Format: "GBSCBilling.<YEAR>-<MONTH>.xlsx"
        ba_filename = os.path.basename(billing_aggregate_file)
        regexp = re.match("GBSCBilling\.([0-9]{4})-([0-9]{2})", ba_filename)

        if regexp is None:
            print >> sys.stderr, "Can't find year-month in %s...skipping" % ba_filename
            continue

        monthbilled = "%s-%s" % (regexp.group(1), regexp.group(2))

        months_billed.append(monthbilled)

        # Open the BillingAggregate workbook.
        print "Reading BillingAggregate workbook %s." % billing_aggregate_file
        billing_aggregate_wkbk = xlrd.open_workbook(billing_aggregate_file, on_demand=True)

        read_billingaggregate_file(billing_aggregate_wkbk, monthbilled)

        billing_aggregate_wkbk.release_resources()

# Sort the resulting months list.
months_billed = sorted(months_billed)

print "Writing out history workbook %s" % args.output
history_wkbk = xlsxwriter.Workbook(args.output)

# Create formats.
money_format   = history_wkbk.add_format({'font_size': 10, 'align': 'right', 'valign': 'top', 'num_format': '$#,##0.00'})
heading_format = history_wkbk.add_format({'bold': True})

write_history_wkbk(history_wkbk, STORAGE_COLUMN_NAME, months_billed, 0)
write_history_wkbk(history_wkbk, COMPUTE_COLUMN_NAME, months_billed, 1)
write_history_wkbk(history_wkbk, CLOUD_COLUMN_NAME, months_billed, 2)
write_history_wkbk(history_wkbk, CONSULTING_COLUMN_NAME, months_billed, 3)
write_history_wkbk(history_wkbk, TOTAL_CHGS_COLUMN_NAME, months_billed, 4)

history_wkbk.close()
