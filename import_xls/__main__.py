from psycopg2 import connect
from json import load
from xlrd import open_workbook, xldate_as_tuple, XL_CELL_DATE, XL_CELL_TEXT, XL_CELL_NUMBER, XL_CELL_EMPTY
from argparse import ArgumentParser
from datetime import datetime
from pprint import pprint
import re
import sys
import traceback
from string import ascii_lowercase

XL_CELL_LONG = 13

type_lookup = {
	XL_CELL_TEXT:	'varchar(255)',
	XL_CELL_DATE:	'timestamp',
	XL_CELL_NUMBER:	'integer',
	XL_CELL_EMPTY:	'integer',
	XL_CELL_LONG:	'text',
}

parser = ArgumentParser()
parser.add_argument('file', help='.xls/.xlsx file for importing.')
parser.add_argument('-s', nargs='+', metavar='s', help='Sheet number.')
parser.add_argument('-o', metavar='offset', help='Row offset.')
parser.add_argument('-c', metavar='c_offset', help='Column offset.')
parser.add_argument('--noreplace', action='store_true', help='Don\'t ask to replace.')
parser.add_argument('--start', metavar='start', help='Start at sheet.')
args = parser.parse_args()

failures = []

def postgresify(string):
	return re.sub(r'(_+)', '_', re.sub(r'( |-|\.|,|\+|\*|")', '_', unicode(string).encode('ascii', 'ignore').lower()))

# Open the workbook
workbook = open_workbook(args.file)
if args.s:
	worksheets = map(lambda i: workbook.sheet_by_index(int(i)), args.s)
elif args.start:
	worksheets = map(lambda i: workbook.sheet_by_index(i), range(int(args.start), len(workbook.sheets())))
else:
	worksheets = workbook.sheets()

# Iterate through each worksheet
for worksheet in worksheets:
	try:
		statements = []

		table_description = worksheet.name.lower()
		table_name = postgresify(table_description)

		print 'Table name: %s' % table_name

		offset = int(args.o) if args.o else int(raw_input('Offset: ') or 0) 
		col_offset = int(args.c) if args.c else int(raw_input('Column offset: ') or 0)

		columns = [None] * col_offset
		descriptions = [None] * col_offset
		types = [None] * col_offset

		# Determines schema for worksheet
		for j in range(col_offset, worksheet.ncols):
			description = worksheet.cell_value(0, j)
			name = postgresify(description)
			columns.append(name)
			descriptions.append(description)
			col_type = XL_CELL_NUMBER
			for i in range(offset, worksheet.nrows):
				this_col_type = worksheet.cell_type(i, j)
				this_col_val = worksheet.cell_value(i, j)
				if this_col_type == XL_CELL_TEXT and len(this_col_val) > 0:
					if len(this_col_val) > 255:
						col_type = XL_CELL_LONG
						break
					else:
						col_type = XL_CELL_TEXT
				elif this_col_type == XL_CELL_DATE:
					col_type = XL_CELL_DATE
					break
			types.append(col_type)

		# Allow user to rename columns
		while True:
			for column in columns[col_offset:]:
				print column
			find = '' if args.noreplace else raw_input('Enter regex to be replaced (or nothing to continue): ')
			if not find:
				break
			replace = raw_input('Enter string to replace with: ')
			columns[col_offset:] = map(lambda s: re.sub(find, replace, s), columns[col_offset:])

		sql = 'CREATE TABLE "%s" (' % table_name
		declarations = []
		for j in range(col_offset, worksheet.ncols):
			declarations.append('"%s" %s' % (columns[j], type_lookup[types[j]]))
		sql += ', '.join(declarations) + ');'

		statements.append({ 'sql': sql, 'values': () })

		sql = 'COMMENT ON TABLE "%s" IS \'"%s"\';' % (table_name, table_description)

		statements.append({ 'sql': sql, 'values': () })

		for j in range(col_offset, worksheet.ncols):
			statements.append({
				'sql': 'COMMENT ON COLUMN "%s"."%s" IS \'%s\';' % (table_name, columns[j], descriptions[j]),
				'values': ()
			})

		# Iterate through each row in each worksheet
		for i in range(offset, worksheet.nrows):
			row = worksheet.row(i)

			values = []

			# Iterate through each cell on the row
			for j in range(col_offset, worksheet.ncols):
				cell_type = types[j]
				if cell_type == XL_CELL_DATE:
					val = datetime(*xldate_as_tuple(worksheet.cell_value(i, j), 0))
				elif cell_type == XL_CELL_NUMBER:
					val = str(int(worksheet.cell_value(i, j) or 0))
				elif cell_type == XL_CELL_TEXT or cell_type == XL_CELL_LONG:
					val = worksheet.cell_value(i, j)
				else:
					raise 'Unhandled type %d.' % cell_type
				values.append(val)

			sql = 'INSERT INTO "%s" VALUES(%s);' % (table_name, ', '.join(['%s'] * len(values)))
			statements.append({ 'sql': sql, 'values': values })

		print
		with open('credentials.json') as credentials:
			connection = connect(**load(credentials))
			with connection.cursor() as cursor:
				for statement in statements:
					# print statement['sql'] % statement['values']
					cursor.execute(statement['sql'], statement['values'])
			connection.commit()
	except Exception as e:
		print
		print 'import_xls: %s' % e
		traceback.print_tb(sys.exc_info()[2])
		failures.append(worksheet.name)
if failures:
	print 'Failed on sheet(s): %s.' % ', '.join(failures)
