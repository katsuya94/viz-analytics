from psycopg2 import connect
from json import load
from xlrd import open_workbook, xldate_as_tuple, XL_CELL_DATE, XL_CELL_EMPTY
from argparse import ArgumentParser
from datetime import datetime
from pprint import pprint
import re

parser = ArgumentParser()
parser.add_argument('file', help='.xls/.xlsx file for importing.')
parser.add_argument('-s', metavar='s', help='Sheet number.')
parser.add_argument('--schema', action='store_true', help='Import schema.')
parser.add_argument('--data', action='store_true', help='Import data.')
args = parser.parse_args()

with open('credentials.json') as credentials:
	connection = connect(**load(credentials))
	cursor = connection.cursor()

# Open the workbook
workbook = open_workbook(args.file)
worksheets = [int(args.s)] if args.s else workbook.sheet_names()

offset = 2

# Iterate through each worksheet
for worksheet_name in worksheets:
	worksheet = workbook.sheet_by_index(worksheet_name)

	columns = []
	descriptions = []

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

	while True:
		pprint(columns)
		find = raw_input('Enter string to be replaced (or nothing to continue): ')
		if not find:
			break
		replace = raw_input('Enter string to replace with: ')
		columns = map(lambda s: re.sub(find, replace, s), columns)


	# Iterate through each row in each worksheet
	for i in range(offset, worksheet.nrows - 1):
		row = worksheet.row(i)

		print '>%d:' % i

		# Iterate through each cell on the row
		for j in range(0, worksheet.ncols - 1):
			if worksheet.cell_type(i, j) == XL_CELL_DATE:
				val = datetime(*xldate_as_tuple(worksheet.cell_value(i, j), 0)).date()
			else:
				val = worksheet.cell_value(i, j)
			print '>>%s: %s' % (columns[j], val)