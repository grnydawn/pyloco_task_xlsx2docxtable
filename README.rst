..  -*- coding: utf-8 -*-

===================
xlsx2docxtable task
===================

version: 0.1.0

converts Microsoft Excel file to a MS Word file.

'xlsx2docxtable' task converts sheets of Microsoft Excel file to tables in a MS Word file.

Installation
------------

Before installing 'xlsx2docxtable' task, please make sure that 'pyloco' is installed.
Run the following command if you need to install 'pyloco'. ::

    >>> pip install pyloco

Or, if 'pyloco' is already installed, upgrade 'pyloco' with the following command ::

    >>> pip install -U pyloco

To install 'xlsx2docxtable' task, run the following 'pyloco' command.  ::

    >>> pyloco install xlsx2docxtable


Command-line syntax
-------------------

usage: pyloco xlsx2docxtable [-h] [-t type] [-o OUTPUT]
                                [--general-arguments]
                                xlsx docx 

converts Microsoft Excel file to a MS Word file.

positional arguments:
  xlsx                  input xlsx file

  docx                  input docx file

optional arguments:

  -h, --help            show this help message and exit
  -t type, --type type  input file format (default='xlsx')
  -o OUTPUT, --output OUTPUT
                        output file
  --general-arguments   Task-common arguments. Use --verbose to see a list of
                        general arguments

forward output variables:
   data                 output data


Example(s)
----------

Current version of the task assumes that an input Excel file is generated
by 'docxtable2xlsx' from an input Word file.

Follwoing command reads "tables.xlsx" Excel file and "my.docx" MS word file,
and convert sheets in "tables.xlsx" to tables of MS Word in "out.docx". ::

    >>> pyloco xlsx2docxtable tables.xlsx my.docx -o out.docx
    out.docx 

Follwoing command reads tables.csv CSV file instead of Excel file in above example. ::

    >>> pyloco xlsx2docxtable tables.csv my.docx -t csv -o out.docx
    out.docx
