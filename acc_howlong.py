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
import os
import os.path
import sys

from ogetools import accounting

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

#=====
#
# SCRIPT BODY
#
#=====

parser = argparse.ArgumentParser()

parser.add_argument("acc_file")
parser.add_argument("-v", "--verbose", action="store_true",
                    default=False,
                    help='Get real chatty [default = false]')

opts = parser.parse_args()

jobstats = accounting.Accounting(opts.acc_file) #, "QLOGIN")

for job in jobstats.getColumns('job_number', 'ru_wallclock', 'ru_utime', 'ru_stime', 'slots'):
    job_id = job[0]
    ru_wallclock = float(job[1])
    ru_utime     = float(job[2])
    ru_stime     = float(job[3])
    wall_hours = ru_wallclock/60/60

    slots        = int(job[4])

    if ru_wallclock > 0.0:
        print job_id, wall_hours, ru_wallclock, ru_utime, ru_stime, (ru_wallclock - ru_utime - ru_stime)/ru_wallclock*100.0, slots

