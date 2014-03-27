#!/usr/bin/env bash

DIRNAME=`dirname $0`
cd "$DIRNAME"

python ./csv_to_xls.py examples/sheet_*.csv -o examples/output.xls -f

has_xls2txt=$(which xls2txt 2> /dev/null)

if [ -z "$has_xls2txt" ]; then
    echo "! Please install xls2txt to view a full diff."
    echo "! Simple diff will be used on binaries as fallback."

    echo "Diff:"
    diff examples/output_ref.xls examples/output.xls
else
    rm -f examples/ref_output.csv
    rm -f examples/output.csv
    xls2txt -A examples/output_ref.xls > examples/ref_output.csv
    xls2txt -A examples/output.xls     > examples/output.csv

    echo "Diff:"
    diff -u examples/ref_output.csv examples/output.csv

    rm -f examples/ref_output.csv
    rm -f examples/output.csv
fi

rm -f examples/output.xls

