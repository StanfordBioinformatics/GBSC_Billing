#!/usr/bin/env python

#===============================================================================
#
# gen_notifs.py - Generate billing notifications for each PI for month/year.
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
from optparse import OptionParser
import os
import os.path
import sys

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

usage = "%prog [options] FILL IN ARGUMENT PATTERN"
parser = OptionParser(usage=usage)

parser.add_option("-v", "--verbose", action="store_true",
                  default=False,
                  help='Get real chatty [default = false]')

(opts, args) = parser.parse_args()

# Validate remaining args.

