#!/usr/bin/env python3

#===============================================================================
#
# convert_Slurm_to_SGE.py - Converts Slurm Accounting file to SGE Accounting file
#
# ARGS:
#   1st: the Slurm Accounting file to be input
#   2nd: the name of the file in which to put the SGE Accounting output
#
# SWITCHES:
#   --billing_root:    Location of BillingRoot directory (overrides BillingConfig.xlsx)
#                      [default if no BillingRoot in BillingConfig.xlsx or switch given: CWD]
#
# INPUT:
#   Slurm Accounting snapshot file (from command sacct --format=ALL).
#
# OUTPUT:
#   SGE Accounting file
#   Various messages about current processing status to STDOUT.
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
import calendar
import csv
import re
import time
import os.path
import sys

# For SGE and Slurm CSV dialect definitions
import job_accounting_file

# Simulate an "include billing_common.py".
SCRIPT_DIR = os.path.dirname(os.path.realpath(__file__))
exec(compile(open(os.path.join(SCRIPT_DIR, "billing_common.py"), "rb").read(), os.path.join(SCRIPT_DIR, "billing_common.py"), 'exec'))

#=====
#
# CONSTANTS
#
#=====


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

def get_nodes_from_nodelist(nodelist_str):

    #
    # I can split the node list by commas, but some node have suffix lists between square brackets, ALSO delimited by commas.
    #

    # Need to convert commas to semicolons in lists marked by [ ]'s
    match_bracket_lists = re.findall('(\[.*?\]+)', nodelist_str)

    for substr in match_bracket_lists:
        new_substr = substr.replace(',', ';')
        nodelist_str = nodelist_str.replace(substr, new_substr)

    # Now, with the commas only separating the node, we can split the node list by commas (still need to unpack square bracket lists).
    node_list_before_unpacking = nodelist_str.split(',')

    node_list = []
    for node in node_list_before_unpacking:

        # Try to break node up into prefix-[suffixes].
        match_prefix_and_bracket_lists = re.search('^(?P<prefix>[^\[]+)\[(?P<suffixes>[^\]]+)\]$', node)

        # If node doesn't match pattern above, add the whole node name.
        if not match_prefix_and_bracket_lists:
            node_list.append(node)
        else:
            match_dict = match_prefix_and_bracket_lists.groupdict()

            prefix = match_dict['prefix']
            suffixes = match_dict['suffixes'].split(';')

            for suffix in suffixes:
                node_list.append(prefix + suffix)

    return node_list

def slurm_time_to_seconds(slurm_time_str):

    # Convert string of form DD-HH:MM:SS to seconds
    elapsed_days_split = slurm_time_str.split('-')
    if len(elapsed_days_split) == 1:
        elapsed_days = 0
        elapsed_hms = elapsed_days_split[0]
    elif len(elapsed_days_split) == 2:
        elapsed_days = int(elapsed_days_split[0])
        elapsed_hms = elapsed_days_split[1]
    else:
        print("Time string of", slurm_time_str, "is malformed.", file=sys.stderr)

    seconds = (elapsed_days * 86400) + sum(int(x) * 60 ** i for i, x in enumerate(reversed(elapsed_hms.split(":"))))

    return seconds


def convert_slurm_file_to_sge_file(slurm_fp, sge_fp):

    # Read in the header line from the Slurm file to use for the DictReader
    header = slurm_fp.readline()
    fieldnames = header.split('|')
    reader = csv.DictReader(slurm_fp,fieldnames=fieldnames,dialect="slurm")

    writer = csv.DictWriter(sge_fp,fieldnames=ACCOUNTING_FIELDS,dialect="sge")
    sge_row = {}  # Create new dictionary for the output row.

    line_num = 0
    for slurm_row in reader:

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

        elapsed_seconds = slurm_time_to_seconds(slurm_row['Elapsed'])
        elapsed_raw_seconds = int(slurm_row['ElapsedRaw'])

        if elapsed_seconds != elapsed_raw_seconds:
            print("Elapsed string of %s does not equal ElapsedRaw value of %d." % (slurm_row['Elapsed'], elapsed_raw_seconds), file=sys.stderr)

        sge_row['ru_wallclock'] = elapsed_seconds

        sge_row['project'] = slurm_row['WCKey']
        sge_row['department'] = "NoDept"
        sge_row['granted_pe'] = "NoPE"
        sge_row['slots'] = slurm_row['NCPUS']

        sge_row['cpu'] = slurm_row['TotalCPU']

        if slurm_row['MaxDiskRead'] != '':
            sge_row['io'] = int(slurm_row['MaxDiskRead'])
        if slurm_row['MaxDiskWrite'] != '':
            sge_row['io'] += int(slurm_row['MaxDiskWrite'])

        if slurm_row['ReqGRES'] == '':
            sge_row['category'] = slurm_row['ReqTRES']
        elif slurm_row['ReqTRES'] == '':
            sge_row['category'] = slurm_row['ReqGRES']
        else:
            sge_row['category'] = "%s;%s" % (slurm_row['ReqTRES'], slurm_row['ReqGRES'])

        sge_row['max_vmem'] = slurm_row['MaxVMSize']

        # Output row to SGE file.
        writer.writerow(sge_row)

        line_num += 1
        if line_num % 10000 == 0:
            sys.stderr.write('.')
            sys.stderr.flush()

    print(file=sys.stderr)

#=====
#
# SCRIPT BODY
#
#=====

parser = argparse.ArgumentParser()

parser.add_argument("--slurm_accounting_file",
                    default=None,
                    help='The Slurm accounting file to read [default = stdin]')
parser.add_argument("--sge_accounting_file",
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
print("Slurm --> SGE Conversion arguments:", file=sys.stderr)
print("  Slurm accounting file: %s" % slurm_accounting_file, file=sys.stderr)
print("  SGE accounting file: %s" % sge_accounting_file, file=sys.stderr)

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

print("Conversion complete!", file=sys.stderr)
