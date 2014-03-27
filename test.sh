#!/usr/bin/env bash

DIRNAME=`dirname $0`
cd "$DIRNAME"

cd examples
rm -f output_ref.csv output.csv output.xls

python ../csv_to_xls.py sheet_*.csv -o output.xls

has_xls2txt=$(which xls2txt 2> /dev/null)

if [ -z "$has_xls2txt" ]; then
    echo "! Please install xls2txt to view a full diff."
    echo "! Simple diff will be used on binaries as fallback."

    diff output_ref.xls output.xls
else
    xls2txt -A output_ref.xls > output_ref.csv
    xls2txt -A output.xls     > output.csv

    diff -u output_ref.csv output.csv
fi

rm -f output_ref.csv output.csv output.xls

