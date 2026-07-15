import sqlite3

conn = sqlite3.connect("database.db")
cursor = conn.cursor()

try:
    cursor.execute("""
        ALTER TABLE employees
        ADD COLUMN deleted_date TEXT
    """)
    print("deleted_date column added")
except:
    print("Column already exists")

conn.commit()
conn.close()