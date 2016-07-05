#!/usr/bin/python
# -*- coding: utf-8 -*-

from __future__ import with_statement

from setuptools import setup


with open('VERSION') as f:
    VERSION = f.read().rstrip()

with open('README.rst') as f:
    LONG_DESCRIPTION = f.read()

with open('LICENSE') as f:
    LICENSE = f.read()

setup(
    name='csv2xls',
    version=VERSION,
    author='Alex Prengère',
    author_email='alexprengere@gmail.com',
    url='https://github.com/alexprengere/csv2xls',
    description='Put together some CSV files into a single Excel file, in different sheets.',
    long_description=LONG_DESCRIPTION,
    license=LICENSE,
    #
    # Manage standalone scripts
    entry_points={
        'console_scripts': [
            'csv2xls = csv_to_xls:main'
        ]
    },
    py_modules=[
        'csv_to_xls'
    ],
    install_requires=[
        'xlwt==0.7.5',
    ],
)
