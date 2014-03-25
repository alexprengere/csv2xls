CSV2Excel
=========

Join together some CSV files into a single Excel file.

Prerequisites
-------------
This is needed:
```bash
pip install --user xlwt
```

Usage
-----
```bash
./csv_to_excel.py

usage: csv_to_excel.py [-h] [-o OUTPUT] [-k] [-f] [-c] files [files ...]

Join together some CSV files into a single Excel file.

positional arguments:
  files

optional arguments:
  -h, --help            show this help message and exit
  -o OUTPUT, --output OUTPUT
                        Define output name.
  -k, --keep-prefix     Keep common prefix when building sheet names.
  -f, --force           If output exists, override it.
  -c, --clean           Remove input files afterwards.
```

Example
-------
```bash
./csv_to_excel.py examples/sheet_alpha.csv examples/sheet_beta.csv
Processing       examples/sheet_alpha.csv -> output.xls/alpha
Processing        examples/sheet_beta.csv -> output.xls/beta
```

