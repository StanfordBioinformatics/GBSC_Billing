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
import calendar
import csv
import time
import os.path
import sys


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
    lineterminator = '\n'
    quotechar = '"'
    quoting = csv.QUOTE_MINIMAL
    skipinitialspace = True
    strict = True
csv.register_dialect("slurm",SlurmDialect)

class SGEDialect(csv.Dialect):

    delimiter = ':'
    doublequote = False
    escapechar = '\\'
    lineterminator = '\n'
    quotechar = '"'
    quoting = csv.QUOTE_MINIMAL
    skipinitialspace = True
    strict = True
csv.register_dialect("sge",SGEDialect)

#=====
#
# GLOBALS
#
#=====

# In billing_common.py
global ACCOUNTING_FIELDS

#=====
#
# FUNCTIONS
#
#=====

# In billing_common.py

def convert_slurm_file_to_sge_file(slurm_fp, sge_fp):

    # Read in the header line from the Slurm file to use for the DictReader
    header = slurm_fp.readline()
    fieldnames = header.split('|')
    reader = csv.DictReader(slurm_fp,fieldnames=fieldnames,dialect="slurm")

    writer = csv.DictWriter(sge_fp,fieldnames=ACCOUNTING_FIELDS,dialect="sge")
    sge_row = {}  # Create new dictionary for the output row.

    line_num = 0
    for slurm_row in reader:

        try:
            sge_row.clear()

            sge_row['qname'] = slurm_row['Partition']
            sge_row['hostname'] = slurm_row['NodeList']
            sge_row['group'] = slurm_row['Group']
            sge_row['owner'] = slurm_row['User']
            sge_row['job_name'] = slurm_row['JobName']
            sge_row['job_number'] = slurm_row['JobIDRaw']
            sge_row['account'] = slurm_row['Account']

            sge_row['submission_time'] = calendar.timegm(time.strptime(slurm_row['Submit'],"%Y-%m-%dT%H:%M:%S"))
            sge_row['start_time'] = calendar.timegm(time.strptime(slurm_row['Start'],"%Y-%m-%dT%H:%M:%S"))
            sge_row['end_time'] = calendar.timegm(time.strptime(slurm_row['End'],"%Y-%m-%dT%H:%M:%S"))

            sge_row['failed'] = 0  # TODO: convert Slurm states to SGE failed states

            (return_value, signal) = slurm_row['ExitCode'].split(':')
            if signal == 0:
                sge_row['exit_status'] = int(return_value)
            else:
                sge_row['exit_status'] = 128 + int(signal)

            # Convert Elapsed of form DD-HH:MM:SS to seconds
            elapsed_days_split = slurm_row['Elapsed'].split('-')
            if len(elapsed_days_split) == 1:
                elapsed_days = 0
                elapsed_hms  = elapsed_days_split[0]
            elif len(elapsed_days_split) == 2:
                elapsed_days = int(elapsed_days_split[0])
                elapsed_hms  = elapsed_days_split[1]
            else:
                print >> sys.stderr, "Elapsed time of", slurm_row['Elapsed'], "is malformed."

            elapsed_seconds = (elapsed_days * 86400) + sum(int(x) * 60 ** i for i,x in enumerate(reversed(elapsed_hms.split(":"))))

            sge_row['ru_wallclock'] = elapsed_seconds

            sge_row['project'] = slurm_row['WCKey']
            sge_row['department'] = "NoDept"
            sge_row['granted_pe'] = "NoPE"
            sge_row['slots'] = slurm_row['NCPUS']

            sge_row['cpu'] = slurm_row['CPUTimeRAW']

            if slurm_row['MaxDiskRead'] != '':
                sge_row['io'] = int(slurm_row['MaxDiskRead'])
            if slurm_row['MaxDiskWrite'] != '':
                sge_row['io'] += int(slurm_row['MaxDiskWrite'])
            sge_row['category'] = slurm_row['ReqGRES']

            sge_row['max_vmem'] = slurm_row['MaxVMSize']
        except ValueError:
            print sge_row

        # Output row to SGE file.
        writer.writerow(sge_row)

        line_num += 1
        if line_num % 10000 == 0:
            print >> sys.stderr, ".",

    print >> sys.stderr

#=====
#
# SCRIPT BODY
#
#=====

parser = argparse.ArgumentParser()

parser.add_argument("slurm_accounting_file",
                    default=None,
                    help='The Slurm accounting file to read [default = stdin]')
parser.add_argument("sge_accounting_file",
                    default=None,
                    help='The SGE accounting file to output to [default = stdout]')
parser.add_argument("-r", "--billing_root",
                    default=None,
                    help='The Billing Root directory [default = None]')
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
else:
    # Use the current directory.
    billing_root = os.getcwd()

# Get absolute path for billing_root directory.
billing_root = os.path.abspath(billing_root)

# Use switch arg for accounting_file if present, else use file in BillingRoot.
if args.slurm_accounting_file is not None:
    slurm_accounting_file = os.path.abspath(args.slurm_accounting_file)
else:
    slurm_accounting_file = "STDIN"

if args.sge_accounting_file is not None:
    sge_accounting_file = os.path.abspath(args.sge_accounting_file)
else:
    sge_accounting_file = "STDOUT"
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
