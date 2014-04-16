import psycopg2, json, csv
from argparse import ArgumentParser

parser = ArgumentParser()
parser.add_argument('file', help='CSV file for importing.')
parser.add_argument('-s', '--schema', action='store_true', help='Import schema.')
parser.add_argument('-d', '--data', action='store_true', help='Import data.')
args = parser.parse_args()

print args

with open('credentials.json') as credentials:
	connection = psycopg2.connect(**json.load(credentials))
	cursor = connection.cursor()

cursor.execute("SELECT * FROM test;")
print cursor.fetchone()