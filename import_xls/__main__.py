from psycopg2 import connect
from json import load
from xlrd import open_workbook, xldate_as_tuple, XL_CELL_DATE, XL_CELL_TEXT, XL_CELL_NUMBER, XL_CELL_EMPTY
from argparse import ArgumentParser
from datetime import datetime
from pprint import pprint
import re
from string import ascii_lowercase

type_lookup = {
	XL_CELL_TEXT: 'varchar(255)',
	XL_CELL_DATE: 'date',
	XL_CELL_NUMBER: 'integer',
	XL_CELL_EMPTY: 'integer',
}

parser = ArgumentParser()
parser.add_argument('file', help='.xls/.xlsx file for importing.')
parser.add_argument('-s', metavar='s', help='Sheet number.')
args = parser.parse_args()

statements = []

# Open the workbook
workbook = open_workbook(args.file)
worksheets = [workbook.sheet_by_index(int(args.s))] if args.s else workbook.sheets()

# Iterate through each worksheet
for worksheet in worksheets:

	table_name = re.sub(r'( |-)', '_', str(worksheet.name.lower()))

	offset = int(raw_input('Offset: ') or 0)
	col_offset = int(raw_input('Column offset: ') or 0)

	columns = [None] * col_offset
	descriptions = [None] * col_offset
	types = [None] * col_offset

	# Determines schema for worksheet
	for j in range(col_offset, worksheet.ncols - 1):
		description = worksheet.cell_value(0, j)
		name = re.sub(r'( |-)', '_', str(description).lower())
		columns.append(name)
		descriptions.append(description)
		col_type = worksheet.cell_type(offset, j)
		if col_type == XL_CELL_TEXT and worksheet.cell_value(offset, j) == '':
			col_type  = XL_CELL_NUMBER
		types.append(col_type)

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
	for i in range(col_offset, worksheet.ncols - 1):
		declarations.append('%s %s' % (columns[i], type_lookup[types[i]]))
	sql += ', '.join(declarations) + ');'

	statements.append(sql)

	# Iterate through each row in each worksheet
	for i in range(offset, worksheet.nrows - 1):
		row = worksheet.row(i)

		values = []

		# Iterate through each cell on the row
		for j in range(col_offset, worksheet.ncols - 1):
			cell_type = types[j]
			if cell_type == XL_CELL_DATE:
				val = 'date \'%s\'' % datetime(*xldate_as_tuple(worksheet.cell_value(i, j), 0)).date()
			elif cell_type == XL_CELL_NUMBER:
				val = str(int(worksheet.cell_value(i, j) or 0))
			elif cell_type == XL_CELL_TEXT:
				val = '\'%s\'' % worksheet.cell_value(i, j)
			else:
				raise 'Unhandled type %d.' % cell_type
			values.append(val)

		sql = 'INSERT INTO %s VALUES(%s);' % (table_name, ', '.join(values))
		statements.append(sql)

with open('credentials.json') as credentials:
	connection = connect(**load(credentials))
	with connection.cursor() as cursor:
		for statement in statements:
			print 'SQL < %s' % statement
			cursor.execute(statement)
	connection.commit()