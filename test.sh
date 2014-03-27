#!/usr/bin/env bash

DIRNAME=`dirname $0`
cd "$DIRNAME"

python ./csv_to_xls.py examples/sheet_alpha.csv examples/sheet_beta.csv -o examples/output.xls
