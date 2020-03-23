#!/bin/python3

import os
import csv
import sys
import re
import argparse
from xlsxwriter.workbook import Workbook

#############################################################################
############################# Reading Arguments #############################
#############################################################################
parser = argparse.ArgumentParser(description="The script will search through the directories looking for CSV files and combined them into one Excel document. The CSV files must start with the following columns: Plugin ID | CVE | CVSS | Risk")

ioGroup = parser.add_argument_group('Input/Output Settings')
ioGroup.add_argument('-d', "--directory", dest="basedirectory", metavar="Directory", help="The root directory for the CSV file search. Default is the current working directory.", default=os.path.dirname(os.path.realpath(__file__)))
ioGroup.add_argument('-o', "--output", dest="outputFile", metavar="OutputFile", help="The filename to save the resulting Excel document. Default is the specified directory (-d) with the filename result.xlsx", default="result.xlsx")

settingsGroup = parser.add_argument_group('General Settings')
settingsGroup.add_argument('-s', "--skip-confirmation", action="store_true", help="Do not ask, start the processing directly")
settingsGroup.add_argument('-f', "--force-overwrite", dest="override", action="store_true", help="Overwrite existing output files without asking")
settingsGroup.add_argument('-G', "--dont-group", action='store_true', help="Do not group rows with the same Plugin ID and Host")
settingsGroup.add_argument('-t', "--total-findings", action='store_true', help="Show total amount of each risk level, even when grouping")
settingsGroup.add_argument("--disable-overview", dest="disableOverview", action='store_true', help="Do create an overview on the first worksheet")
settingsGroup.add_argument('-H', "--hide-empty", action='store_true', help="Hide worksheets with empty tables")

filterGroup = parser.add_argument_group('Filter Settings')
filterGroup.add_argument('-A', "--include-all", action="store_true", help="Include every risk level and disable the CVSS filtering")
filterGroup.add_argument('-r', "--include-risks", default="Critical,High,Medium", help="Select the risk levels from the .csv file that should be shown in the table (comma separated, case sensitive!) -> Default: Critical,High,Medium")
filterGroup.add_argument("--no-cvss-filter", action='store_true', help="Dont hide entries with misssing CVSS score or CVSS score of 0")
filterGroup.add_argument("--optional-findings", action='store_true', help="If there are no findings of the included risk levels, show everything else")

designGroup = parser.add_argument_group('Design Settings')
designGroup.add_argument("--table-style", default="Table Style Medium 15", help="Set the style of the table using the Excel templates. Default. 'Table Style Medium 15'")
designGroup.add_argument("--disable-highlighting", action='store_true', help="Do not highlight critical and high findings")
