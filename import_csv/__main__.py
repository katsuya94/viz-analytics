import psycopg2
import json

connection = psycopg2.connect(**json.load(open('credentials.json')))
cursor = connection.cursor()

cursor.execute("SELECT * FROM test;")
print cursor.fetchone()