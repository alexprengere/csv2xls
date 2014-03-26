csv2xls
=======

Put together some CSV files into a single Excel file, in different sheets.

Installation
------------

You can install directly after cloning:

.. code-block:: bash

 $ python setup.py install --user

Or use the Python package:

.. code-block:: bash

 $ pip install --user csv2xls

Dependency
----------
Outside the standard library, the *xlwt* package is needed, and should be
automatically installed with setuptools. Otherwise:

.. code-block:: bash

 $ pip install --user xlwt

Usage
-----

.. code-block:: bash

 $ ./csv_to_xls.py -h
 usage: csv_to_xls.py [-h] [-o OUTPUT] [-k] [-f] [-c] [-d DATE_FORMAT]
                      files [files ...]

 Put together some CSV files into a single Excel file. Basic types are infered
 automatically.

 positional arguments:
   files

 optional arguments:
   -h, --help            show this help message and exit
   -o OUTPUT, --output OUTPUT
                         Define name for output Excel file. Default is
                         output.xls.
   -k, --keep-prefix     Keep common prefix when building sheet names.
   -f, --force           If output already exists, override it.
   -c, --clean           Delete input files after successfully creating the
                         Excel file.
   -d DATE_FORMAT, --date-format DATE_FORMAT
                         Change date format used during date type inference.
                         Default is %Y-%m-%d.

 Example: ./csv_to_xls.py examples/sheet_alpha.csv examples/sheet_beta.csv

Example
-------

.. code-block:: bash

 $ ./csv_to_xls.py examples/sheet_alpha.csv examples/sheet_beta.csv
 Processing       examples/sheet_alpha.csv -> output.xls/alpha
 Processing        examples/sheet_beta.csv -> output.xls/beta

