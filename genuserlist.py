#!/usr/bin/env python

from optparse import OptionParser
import pwd
import grp

import xlrd

# List of non-lab group names.
non_lab_group_names = [ "scgpm-informatics_vendors",
#                        "gbsc-pacbio",
#                        "training-camp"
]

# Mapping from group name to PI.
# group_name_to_pi = { 
#     "scg-admin" : "Somalee",
#     "scgpm-informatics_ashley" : "Ashley",
#     "scgpm-informatics_assimes" : "Assimes",
#     "mb-lab" : "Barna",
#     "bat-lab" : "Batzoglou",
#     "scgpm-informatics_blau" : "Blau",
#     "ab-lab" : "Brunet",
#     "cb-lab" : "Bustamante",
#     "butte_lab" : "Butte",
#     "mc-lab" : "Cherry",
#     "mfeldman" : "Feldman",
#     "mf-lab" : "Fuller",
#     "ag-lab" : "Gitler",
#     "wg-lab" : "Greenleaf",
#     "hj-lab" : "Hanlee Ji",
#     "scgpm-informatics_kundaje" : "Kundaje",
#     "bl-lab" : "Billy Li",
#     "sm-lab" : "Montgomery",
#     "dp-lab" : "Petrov",
#     "pol-lab" : "Pollack",
#     "jp-lab" : "Pringle",
#     "scgpm-informatics_pritchard" : "Pritchard",
#     "scgpm-informatics_tomq1lab" : "Quertermous",
#     "js-lab" : "Sage",
#     "as-lab" : "Sidow",
#     "ms-lab" : "Snyder",
#     "ht-lab" : "Tang",
#     "scgpm-informatics_urban" : "Urban",
#     "mw-lab" : "Winslow",
#     "scgpm-informatics_wong" : "Wing Wong",
#     "training-camp" : "Genetics Training Camp",
#     "scgpm-informatics_rosenberg" : "Rosenberg"
# }

# Mapping from group name to group member list.
group_members = dict()

# Set of all users we are interested in.
all_users = set()

# Set of users in at least one lab.
all_lab_members = set()

# Set of users in at least one non-lab group.
all_non_lab_members = set()

# Set of all users in at least one lab or non-lab group.
all_group_members = set()


# This method takes in an xlrd Sheet object and a column name,
# and returns all the values from that column.
def sheet_get_named_column(sheet, col_name):

    header_row = sheet.row_values(0)

    for idx in range(len(header_row)):
        if header_row[idx] == col_name:
           col_name_idx = idx
           break
    else:
        return None

    return sheet.col_values(col_name_idx,start_rowx=1)

def read_billing_conf_db(db_workbook):

   pi_sheet = db_workbook.sheet_by_name("PIs")

   group_names   = sheet_get_named_column(pi_sheet, "Group Name")
   pi_last_names = sheet_get_named_column(pi_sheet, "PI Last Name")
   pi_tags       = sheet_get_named_column(pi_sheet, "PI Tag")

   group_name_to_pi     = dict(zip(group_names, pi_last_names))
   group_name_to_pi_tag = dict(zip(group_names, pi_tags))

   return (group_name_to_pi, group_name_to_pi_tag)


usage = "%prog [options]"
parser = OptionParser(usage=usage)

parser.add_option("--lab_users", action="store_true",
                  default=False,
                  help='Show lab membership [default = false]')
parser.add_option("--non_lab_users", action="store_true",
                  default=False,
                  help='Show non-lab membership [default = false]')
parser.add_option("--users_no_lab", action="store_true",
                  default=False,
                  help='Show users with no lab [default = false]')
parser.add_option("--multi_lab", action="store_true",
                  default=False,
                  help='Show users with more than 1 lab [default = false]')

(options, args) = parser.parse_args()

#
# Read in the Billing Configuration DB .xlsx file.
#
billing_conf_db = args[0]
db_workbook = xlrd.open_workbook(billing_conf_db)
(group_name_to_pi, group_name_to_pi_tag) = read_billing_conf_db(db_workbook)

# Set of all "lab" group names.
all_lab_groups = set(group_name_to_pi.keys())

pi_to_group_name = dict(zip(group_name_to_pi.values(), group_name_to_pi.keys()))
pi_tag_to_group_name = dict(zip(group_name_to_pi_tag.values(), group_name_to_pi_tag.keys()))

#
# Scan passwd DB to find primary groups for all users.
#
users = pwd.getpwall()
for user in users:
    if user.pw_uid >= 500:
        group_name = grp.getgrgid(user.pw_gid).gr_name
        if group_members.get(group_name) is None:
            group_members[group_name] = [user.pw_name]
        else:
            group_members[group_name].append(user.pw_name)
        
        # Add this user to list of all users.
        all_users.add(user.pw_name)

#
# Add users from group member lists.
#
group_names = group_name_to_pi.keys() + non_lab_group_names

for group_name in group_names:
    gr_db_entry = grp.getgrnam(group_name)
    if group_members.get(group_name) is None:
        group_members[group_name] = gr_db_entry.gr_mem
    else:
        group_members[group_name].extend(gr_db_entry.gr_mem)

#
# Create set of users in at least one lab group.
#
for group_name in all_lab_groups:
    all_lab_members.update(group_members[group_name])
    all_group_members.update(group_members[group_name])

for group_name in non_lab_group_names:
    all_non_lab_members.update(group_members[group_name])
    all_group_members.update(group_members[group_name])

# Find the users not in any lab or non-lab group.
not_in_lab_users = all_users - all_group_members

if options.users_no_lab:
    print
    print "Users not in any lab:"
    for user in not_in_lab_users:
        print user

#
# Compute groups for each user.
#
groups_per_user = dict()

for group_name in group_members.iterkeys():
    for user in group_members[group_name]:
        if groups_per_user.get(user) is None:
            groups_per_user[user] = [group_name]
        else:
            groups_per_user[user].append(group_name)

if options.multi_lab:
    print
    print "Users in more than one lab group:"
    for user in groups_per_user.iterkeys():
        lab_groups_for_user = []
        for group_name in groups_per_user[user]:
            if group_name in all_lab_groups:
                lab_groups_for_user.append(group_name)
        if len(lab_groups_for_user) > 1:  
            print user,
            for group in lab_groups_for_user:
                print group,
            print

if options.lab_users:
    #
    # List users per group
    #
    print
    print "Users per lab group (%d total users):" % (len(all_lab_members))
    print
    print "PI Last Name\tUsername\tEmail\tFull Name"
    for pi in sorted(pi_to_group_name):
        group_name = pi_to_group_name[pi]
        if len(group_members[group_name]) > 0:
            if group_name in group_name_to_pi:
                pi = group_name_to_pi[group_name]
            else:
                pi = group_name

            for member in sorted(group_members[group_name]):
                try:
                    fullname = pwd.getpwnam(member).pw_gecos
                except KeyError:
                    fullname = "NO ACCOUNT"

                if fullname == '':
                    fullname = "NO NAME"

                print "%s\t%s\t%s@stanford.edu\t%s" % (pi, member, member, fullname)

if options.non_lab_users:
    #
    # List users per group
    #
    print
    print "Users per non-lab group (%d total users):" % (len(all_non_lab_members))
    print
    for group_name in sorted(non_lab_group_names):
        if len(group_members[group_name]) > 0:
            if group_name in group_name_to_pi:
                pi = group_name_to_pi[group_name]
            else:
                pi = group_name
            print "%s (%s) [%d members]:" % (pi, group_name, len(group_members[group_name]))
            for member in sorted(group_members[group_name]):
                try:
                    fullname = pwd.getpwnam(member).pw_gecos
                except KeyError:
                    fullname = "NO ACCOUNT"

                if fullname == '':
                    fullname = "NO NAME"

                print "%s\t%s\t%s@stanford.edu\t%s" % (pi, member, member, fullname)
            print
