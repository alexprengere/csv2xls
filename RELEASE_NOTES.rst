=============
Release notes
=============

V0
==

+ 0.4 :

    + PEP8 compliance by @lowks
    + Add possibility to configure delimiter and quotechar with ``-d`` and ``-q``
    + ``-d`` is now ``-D`` to configure date format for type inference
    + fix bug in option name

+ 0.3 :

    + Allow multiple same sheet in input
    + Make force the default mode, remove the option (``-f``).
    + Add a mode with no type inference, raw dumping of data (``-n``)
    + Prevent the ``ValueError`` from number of rows > 65535
    + Add ``":"`` to forbidden chars in sheet names.
    + Fix bug where only one sheet and keep_prefix is activated

+ 0.2 :

    + add tests
    + add *float/int/date* type inference

+ 0.1 :

    + polish CLI
    + first working version using *xlwt*

