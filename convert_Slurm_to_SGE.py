#!/usr/bin/env python

#===============================================================================
#
# gen_details.py - Generate billing details for month/year.
#
# ARGS:
#   1st: the BillingConfig spreadsheet.
#
# SWITCHES:
#   --accounting_file: Location of accounting file (overrides BillingConfig.xlsx)
#   --billing_root:    Location of BillingRoot directory (overrides BillingConfig.xlsx)
#                      [default if no BillingRoot in BillingConfig.xlsx or switch given: CWD]
#   --year:            Year of snapshot requested. [Default is this year]
#   --month:           Month of snapshot requested. [Default is last month]
#   --no_storage:      Don't run the storage calculations.
#   --no_usage:        Don't run the storage usage calculations (only the quotas).
#   --no_computing:    Don't run the computing calculations.
#   --no_consulting:   Don't run the consulting calculations.
#   --all_jobs_billable: Consider all jobs to be billable. [default=False]
#   --ignore_job_timestamps: Ignore timestamps in job and allow jobs not in month selected [default=False]
#
# INPUT:
#   BillingConfig spreadsheet.
#   SGE Accounting snapshot file (from snapshot_accounting.py).
#     - Expected in BillingRoot/<year>/<month>/SGEAccounting.<year>-<month>.xlsx
#
# OUTPUT:
#   BillingDetails spreadsheet in BillingRoot/<year>/<month>/BillingDetails.<year>-<month>.xlsx
#   Various messages about current processing status to STDOUT.
#
# ASSUMPTIONS:
#   Depends on xlrd and xlsxwriter modules.
#   The input spreadsheet has been certified by check_config.py.
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
import codecs
import collections
import csv
import datetime
import locale
import os
import os.path
import subprocess
import sys
import time

import xlrd
import xlsxwriter

# Simulate an "include billing_common.py".
SCRIPT_DIR = os.path.dirname(os.path.realpath(__file__))
execfile(os.path.join(SCRIPT_DIR, "billing_common.py"))

#=====
#
# CONSTANTS
#
#=====

#=====
#
# CLASSES
#
#=====
class SlurmDialect(csv.Dialect):

    delimiter = '|'
    doublequote = False
    escapechar = '\\'
    quoting = csv.QUOTE_MINIMAL
    skipinitialspace = True
    strict = True

#=====
#
# GLOBALS
#
#=====

#=====
#
# FUNCTIONS
#
#=====

# In billing_common.py

def convert_slurm_file_to_sge_file(slurm_fp, sge_fp):


    for line in slurm_fp:


#=====
#
# SCRIPT BODY
#
#=====

parser = argparse.ArgumentParser()

parser.add_argument("slurm_accounting_file",
                    default=None,
                    help='The Slurm accounting file to read [default = stdin]')
parser.add_argument("-o", "sge_accounting_file",
                    default=None,
                    help='The SGE accounting file to output to [default = stdout]')
parser.add_argument("-v", "--verbose", action="store_true",
                    default=False,
                    help='Get real chatty [default = false]')

args = parser.parse_args()

#
# Process arguments.
#

# Override billing_root with switch args, if present.
if args.billing_root is not None:
    billing_root = args.billing_root
# If we still don't have a billing root dir, use the current directory.
if billing_root is None:
    billing_root = os.getcwd()

# Get absolute path for billing_root directory.
billing_root = os.path.abspath(billing_root)

# Use switch arg for accounting_file if present, else use file in BillingRoot.
if args.slurm_accounting_file is not None:
    slurm_accounting_file = args.slurm_accounting_file
else:
    slurm_accounting_file = "STDIN"

if args.sge_accounting_file is not None
    sge_accounting_file = args.sge_accounting_file
else:
    sge_accounting_file = "STDOUT"

# Get absolute path for accounting files.
slurm_accounting_file = os.path.abspath(slurm_accounting_file)
sge_accounting_file = os.path.abspath(sge_accounting_file)

#
# Output the state of arguments.
#
print >> sys.stderr, "Slurm --> SGE Conversion arguments:"
print >> sys.stderr, "  Slurm accounting file: %s" % slurm_accounting_file
print >> sys.stderr, "  SGE accounting file: %s" % sge_accounting_file

#
# Open the two files
#
if slurm_accounting_file == "STDIN":
    slurm_accounting_fp = sys.stdin
else:
    slurm_accounting_fp = open(slurm_accounting_file, "r")

if sge_accounting_file == "STDOUT":
    sge_accounting_fp = sys.stdout
else:
    sge_accounting_fp = open(sge_accounting_file, "w")

convert_slurm_file_to_sge_file(slurm_accounting_fp, sge_accounting_fp)

print >> sys.stderr, "Conversion complete!"
