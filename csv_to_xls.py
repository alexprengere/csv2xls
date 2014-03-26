#!/usr/bin/env python

"""
Put together some CSV files into a single Excel file, in different sheets.
"""

from __future__ import with_statement

import os, sys
import os.path as op
from datetime import datetime
from collections import defaultdict
import csv

import xlwt

DEF_DATE_FORMAT = "%Y-%m-%d"


def sanitize(name):
    """xlwt does not allow long sheet names.
    """
    for c in '/', '?':
        # For '/', we do not print
        if c in name and c != '/':
            print("! Sheet names cannot contain '{0}', replacing in {1}".format(c, name))
        name = name.replace(c, '_')

    limit = 28
    if len(name) > limit:
        print("! Sheet name too long. Trimming {0} to {1}".format(name, name[:limit]))
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
    for f, sheet_name in sheet_names.items():
        count_sheet_names[sheet_name].append(f)

    for sheet_name, list_files in count_sheet_names.items():
        if len(list_files) > 1:
            # Duplicates here
            for i, f in enumerate(list_files, start=1):
                sheet_names[f] = '{0}_{1}'.format(sheet_name, i)
                print("! To avoid duplicated sheet names, renaming {0} to {1}".format(sheet_name, sheet_names[f]))

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


# XFS style for date format
DATE_FORMAT_STYLE = xlwt.XFStyle()
DATE_FORMAT_STYLE.num_format_str = 'M/D/YY'

def write_to_sheet(sheet, row_nb, col_nb, v, date_format):
    """Custom sheet writer with type inference.
    """
    if is_int(v):
        sheet.write(row_nb, col_nb, int(v))

    elif is_float(v):
        sheet.write(row_nb, col_nb, float(v))

    elif is_date(v, date_format):
        sheet.write(row_nb, col_nb,
                    datetime.strptime(v, date_format),
                    DATE_FORMAT_STYLE)
    else:
        sheet.write(row_nb, col_nb, v)


def add_to_sheet(sheet, fl, date_format):
    """Add filelike content to sheet.
    """
    for row_nb, row in enumerate(csv.reader(fl, delimiter=',', quotechar='"')):

        for col_nb, v in enumerate(row):
            # Type inference hidden here
            write_to_sheet(sheet, row_nb, col_nb, v, date_format)


def create_excel_file(files, output, keep_prefix, force, date_format, clean):
    """Main function creating the excel file.
    """
    if not output.endswith(".xls") and not output.endswith(".xlsx"):
        print("! Output name should end with .xls[x] extension, got:")
        print("{0:^40}".format(output))
        return

    if op.exists(output) and not force:
        print("! Output already exists: {0}".format(output))
        return

    # THE Excel book ;)
    book = xlwt.Workbook()

    for f, sheet_name in sorted(build_sheet_names(files, keep_prefix).items(),
                                key=lambda t: t[1].lower()):

        print("Processing {0:>30} -> {1}/{2}".format(f, output, sheet_name))

        with open(f) as fl:
            sheet = book.add_sheet(sheet_name)
            add_to_sheet(sheet, fl, date_format)

    book.save(output)

    # Hopefully no exception raised so far
    if clean:
        for f in sorted(files):
            print("Removing {0}".format(f))
            os.unlink(f)


def main():
    """Main.
    """
    import argparse

    parser = argparse.ArgumentParser(description="""
    Put together some CSV files into a single Excel file.
    Basic types are infered automatically.
    """)

    parser.add_argument("files", nargs='+')

    parser.add_argument("-o", "--output",
        help="""
        Define name for output Excel file.
        Default is %(default)s.""",
        default="output.xls")

    parser.add_argument("-k", "--keep-prefix",
        help="""
        Keep common prefix when
        building sheet names.
        """,
        action='store_true')

    parser.add_argument("-f", "--force",
        help="""
        If output already exists, override it.
        """,
        action='store_true')

    parser.add_argument("-c", "--clean",
        help="""
        Delete input files after successfully creating the Excel file.
        """,
        action='store_true')

    parser.add_argument("-d", "--date-format",
        help="""
        Change date format used during date type
        inference. Default is %(default)s.
        """,
        default=DEF_DATE_FORMAT)

    parser.epilog = """
    Example: {0} examples/sheet_alpha.csv examples/sheet_beta.csv
    """.format(op.basename(sys.argv[0]))

    args = parser.parse_args()

    create_excel_file(args.files,
                      args.output,
                      args.keep_prefix,
                      args.force,
                      args.date_format,
                      args.clean)


if __name__ == "__main__":

    main()

