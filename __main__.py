# -*- utf-8 -*-

#
# Main program entry point for NAV parser.
#
# 1. Scan emails and download NAV table files.
#
# 2. Parse NAV table files and extract useful information (NAV, Unit NAV, Accumulated return, etc)
#
# 3. Fill the report Excel table and send it to related persons.
#
# All rights reserved, 2019
#


""" A main program for NAV email parser."""


from NavReporter import NavReporter


def entry_point():
    conf_file = './NavReporter/config.cfg'
    nav_rpt = NavReporter()
    nav_rpt.run(conf_file)
    
run_fn = entry_point
run_fn()
