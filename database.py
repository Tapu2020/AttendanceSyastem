import sqlite3

conn = sqlite3.connect("database.db")
cursor = conn.cursor()

# Employees Table
cursor.execute("""
CREATE TABLE IF NOT EXISTS employees(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    mobile TEXT,
    aadhaar TEXT,
    joining_date TEXT,
    monthly_salary REAL,
    shift_start TEXT,
    shift_end TEXT,
    grace_minutes INTEGER DEFAULT 15,
    rfid_uid TEXT,
    status TEXT DEFAULT 'Active',
    deleted_date TEXT
)
""")

# Attendance Table
cursor.execute("""
CREATE TABLE IF NOT EXISTS attendance(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    employee_id INTEGER,
    attendance_date TEXT,
    check_in TEXT,
    check_out TEXT,
    working_hours TEXT,
    status TEXT,
    late_minutes INTEGER DEFAULT 0
)
""")

conn.commit()
conn.close()

print("Database Created Successfully")