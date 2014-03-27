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
Outside the standard library, the `xlwt <http://www.python-excel.org/>`_ package is needed, and should be
automatically installed with setuptools. Otherwise:

.. code-block:: bash

 $ pip install --user xlwt

Example
-------

.. code-block:: bash

 $ csv2xls examples/sheet_alpha.csv examples/sheet_beta.csv -o output.xls
 Processing       examples/sheet_alpha.csv -> output.xls/alpha
 Processing        examples/sheet_beta.csv -> output.xls/beta

Usage
-----

.. code-block:: bash

 $ csv2xls -h
 usage: csv2xls [-h] [-o OUTPUT] [-k] [-f] [-c] [-d DATE_FORMAT]
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

 Example: csv2xls examples/sheet_alpha.csv examples/sheet_beta.csv

Tests
-----
To run the tests, you must install `xls2txt <https://github.com/hroptatyr/xls2txt>`_:

.. code-block:: bash

 $ git clone https://github.com/hroptatyr/xls2txt.git
 $ cd xls2txt
 $ make
 $ sudo make install

Then run:

.. code-block:: bash

 $ ./tests.sh

