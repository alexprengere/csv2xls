#!/usr/bin/python
# -*- coding: utf-8 -*-

from __future__ import with_statement

import os.path as op
from setuptools import setup

def local(rel_path, root_file=__file__):
    return op.join(op.realpath(op.dirname(root_file)), rel_path)


with open(local('VERSION')) as fl:
    VERSION = fl.read().rstrip()

with open(local('README.rst')) as fl:
    LONG_DESCRIPTION = fl.read()

with open(local('LICENSE')) as fl:
    LICENSE = fl.read()

setup(
    name = 'csv2xls',
    version = VERSION,
    author = 'Alex Prengère',
    author_email = 'alexprengere@gmail.com',
    url = 'https://github.com/alexprengere/csv2xls',
    description = 'Put together some CSV files into a single Excel file, in different sheets.',
    long_description = LONG_DESCRIPTION,
    license = LICENSE,
    #
    # Manage standalone scripts
    entry_points = {
        'console_scripts' : [
            'csv2xls = csv_to_xls:main'
        ]
    },
    py_modules = [
        'csv_to_xls'
    ],
    install_requires = [
        'xlwt==0.7.5',
    ],
)

