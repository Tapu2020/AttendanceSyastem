import sqlite3

conn = sqlite3.connect("database.db")
cursor = conn.cursor()

try:
    cursor.execute(
        "ALTER TABLE attendance ADD COLUMN working_hours TEXT"
    )
except:
    pass

try:
    cursor.execute(
        "ALTER TABLE attendance ADD COLUMN late_minutes INTEGER DEFAULT 0"
    )
except:
    pass

conn.commit()
conn.close()

print("Attendance table upgraded")