#!/usr/bin/env python

"""
Join several CSV files into one single Excel WorkBook.
"""

from __future__ import with_statement

import xlwt
import csv
import os
import os.path as op
from collections import defaultdict


def sanitize(name):
    """xlwt does not allow long sheet names.
    """
    for c in '/', '?':
        # For '/', we do not print
        if c in name and c != '/':
            print "Sheet names cannot contain '{0}', replacing in {1}".format(c, name)
        name = name.replace(c, '_')

    limit = 28
    if len(name) > limit:
        print "Sheet names are limited to {0} characters!".format(limit)
        print "Trimming {0} to {1}".format(name, name[:limit])
    return name[:limit]


def build_sheet_names(files, keep_prefix):
    """Build nice sheet names from file names.

    We trim the common prefix, remove the extension.
    """
    if keep_prefix:
        prefix = ''
    else:
        prefix = op.commonprefix(files)

    # Helper lambdas
    trim_prefix = lambda s: s[len(prefix):]
    trim_extens = lambda s: op.splitext(s)[0]

    # Remove prefix, extension
    sheet_names = {}
    for f in files:
        sheet_names[f] = sanitize(trim_extens(trim_prefix(f)))

    # Handling duplicates
    count_sheet_names = defaultdict(list)
    for f, sheet_name in sheet_names.iteritems():
        count_sheet_names[sheet_name].append(f)

    for sheet_name, list_files in count_sheet_names.iteritems():
        if len(list_files) > 1:
            # Duplicates here
            for i, f in enumerate(list_files, start=1):
                sheet_names[f] = '{0}_{1}'.format(sheet_name, i)
                print "To avoid duplicated sheet names, renaming {0} to {1}".format(sheet_name, sheet_names[f])

    return sheet_names


def add_to_sheet(sheet, fl):
    """Add filelike content to sheet.
    """
    for num, line in enumerate(csv.reader(fl, delimiter=',', quotechar='"')):
        row = sheet.row(num)
        for index, elem in enumerate(line):
            row.write(index, elem)


def create_excel_file(sheet_names, output):
    """Main function creating the excel file.
    """
    book = xlwt.Workbook()

    for f, sheet_name in sorted(sheet_names.iteritems(), key=lambda n: n[1].lower()):
        print "Processing {0:>30} -> {1}/{2}".format(f, output, sheet_name)
        with open(f) as fl:
            sheet = book.add_sheet(sheet_name)
            add_to_sheet(sheet, fl)

    book.save(output)


def main(args):
    """Main.
    """
    if not args.output.endswith(".xls") and not args.output.endswith(".xlsx"):
        print "Output name should end with .xls[x] extension, got:"
        print "{0:^40}".format(args.output)
        exit(1)

    if op.exists(args.output) and not args.force:
        print "Output already exists: {0}".format(args.output)
        exit(1)

    create_excel_file(build_sheet_names(args.files, args.keep_prefix),
                      args.output)

    # Hopefully no exception raised so far
    if args.clean:
        for f in sorted(args.files):
            print "Removing {0}".format(f)
            os.unlink(f)



if __name__ == "__main__":

    import argparse

    parser = argparse.ArgumentParser(description="""
    Join together some CSV files into a single Excel file.
    """)

    parser.add_argument("files", nargs='+')

    parser.add_argument("-o", "--output",
        help="Define output name.",
        default="output.xls")

    parser.add_argument("-k", "--keep-prefix",
        help="""
        Keep common prefix when
        building sheet names.
        """,
        action='store_true')

    parser.add_argument("-f", "--force",
        help="""
        If output exists, override it.
        """,
        action='store_true')

    parser.add_argument("-c", "--clean",
        help="""
        Remove input files afterwards.
        """,
        action='store_true')

    args = parser.parse_args()

    main(args)

