#! python3
'''
Purpose of this file:
Build your excel file on a copy of this file, which imports my default spreadsheet file, spreadsheet_default.py
'''
# Directly importing a module works once. For repeatitive importing, use importlib.reload(fileName)

print('\n  spreadsheet_master.py launched')

import sys
for path in sys.path:
    print('path = ', path)

import importlib
importlib.import_module('spreadsheet_default') # No .py needed, causes 'ModuleNotFoundError'

print('sheetnames = ', wb.sheetnames)

print('\n  spreadsheet_master.py complete')

