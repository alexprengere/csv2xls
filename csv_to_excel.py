#!/usr/bin/env python

"""
Join several CSV files into one single Excel WorkBook.
"""

from __future__ import with_statement

import xlwt
import csv
import os
import os.path as op
from datetime import datetime
from collections import defaultdict

DEF_DATE_FORMAT = "%Y-%m-%d"


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
        print "Name too long! Trimming {0} to {1}".format(name, name[:limit])
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
                print "! To avoid duplicated sheet names, renaming {0} to {1}".format(sheet_name, sheet_names[f])

    return sheet_names


def is_int(s):
    """Type inference when writing in Excel.
    """
    try:
        int(s)
    except ValueError:
        return False
    else:
        return True


def is_float(s):
    """Type inference when writing in Excel.
    """
    try:
        float(s)
    except ValueError:
        return False
    else:
        return True


def is_date(s, date_format):
    """Type inference when writing in Excel.
    """
    try:
        datetime.strptime(s, date_format)
    except ValueError:
        return False
    else:
        return True


def write_to_row(row, index, v, date_format):
    """Custom row writer with type inference.
    """
    if is_int(v):
        row.write(index, int(v))

    elif is_float(v):
        row.write(index, float(v))

    elif is_date(v, date_format):
        # XFS style for date format
        date_format_style = xlwt.XFStyle()
        date_format_style.num_format_str = 'M/D/YY'
        row.write(index,
                  datetime.strptime(v, date_format),
                  date_format_style)
    else:
        row.write(index, v)



def add_to_sheet(sheet, fl, date_format):
    """Add filelike content to sheet.
    """
    for num, line in enumerate(csv.reader(fl, delimiter=',', quotechar='"')):
        row = sheet.row(num)

        for index, v in enumerate(line):
            # Type inference hidden here
            write_to_row(row, index, v, date_format)


def create_excel_file(sheet_names, output, date_format):
    """Main function creating the excel file.
    """
    book = xlwt.Workbook()

    for f, sheet_name in sorted(sheet_names.iteritems(), key=lambda (_, v): v.lower()):
        print "Processing {0:>30} -> {1}/{2}".format(f, output, sheet_name)
        with open(f) as fl:
            sheet = book.add_sheet(sheet_name)
            add_to_sheet(sheet, fl, date_format)

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
                      args.output,
                      args.date_format)

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

    parser.add_argument("-d", "--date-format",
        help="""
        Change date format used for date type
        inference. Default is {0}.
        """.format(DEF_DATE_FORMAT),
        default=DEF_DATE_FORMAT)

    args = parser.parse_args()

    main(args)

