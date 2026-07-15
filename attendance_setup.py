import sqlite3

conn = sqlite3.connect("database.db")
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS attendance(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    employee_id INTEGER,
    attendance_date TEXT,
    check_in TEXT,
    check_out TEXT,
    working_hours TEXT,
    status TEXT
)
""")

conn.commit()
conn.close()

print("Attendance table created")