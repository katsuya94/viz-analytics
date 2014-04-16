from psycopg2 import connect
from json import load
from xlrd import open_workbook, xldate_as_tuple, XL_CELL_DATE, XL_CELL_TEXT, XL_CELL_NUMBER
from argparse import ArgumentParser
from datetime import datetime
from pprint import pprint
import re

type_lookup = {
	XL_CELL_TEXT: 'varchar(255)',
	XL_CELL_DATE: 'date',
	XL_CELL_NUMBER: 'integer'
}

parser = ArgumentParser()
parser.add_argument('file', help='.xls/.xlsx file for importing.')
parser.add_argument('-s', metavar='s', help='Sheet number.')
args = parser.parse_args()

with open('credentials.json') as credentials:
	connection = connect(**load(credentials))
	cursor = connection.cursor()

# Open the workbook
workbook = open_workbook(args.file)
worksheets = [workbook.sheet_by_index(int(args.s))] if args.s else workbook.sheets()

offset = 2

# Iterate through each worksheet
for worksheet in worksheets:

	table_name = re.sub(r'( |-)', '_', str(worksheet.name.lower()))

	columns = []
	descriptions = []
	types = []

	# Determines schema for worksheet
	for j in range(0, worksheet.ncols - 1):
		description = worksheet.cell_value(0, j)
		name = re.sub(r'( |-)', '_', str(description).lower())
		occurrences = columns.count(name)
		if occurrences == 0:
			columns.append(name)
		else:
			columns.append('%s_%d' % (name, occurrences))
		descriptions.append(description)
		types.append(type_lookup[worksheet.cell_type(offset, j)])

	# Allow user to rename columns
	while True:
		pprint(columns)
		find = raw_input('Enter regex to be replaced (or nothing to continue): ')
		if not find:
			break
		replace = raw_input('Enter string to replace with: ')
		columns = map(lambda s: re.sub(find, replace, s), columns)

	sql = 'CREATE TABLE %s (' % table_name
	declarations = []
	for i in range(0, worksheet.ncols - 1):
		declarations.append('%s %s' % (columns[i], types[i]))
	sql += ', '.join(declarations) + ');'

	print sql

	# Iterate through each row in each worksheet
	for i in range(offset, worksheet.nrows - 1):
		row = worksheet.row(i)

		# Iterate through each cell on the row
		for j in range(0, worksheet.ncols - 1):
			if worksheet.cell_type(i, j) == XL_CELL_DATE:
				val = datetime(*xldate_as_tuple(worksheet.cell_value(i, j), 0)).date()
			else:
				val = worksheet.cell_value(i, j)