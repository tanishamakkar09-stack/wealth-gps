import sqlite3

# Create / Connect Database
conn = sqlite3.connect("database.db")

cursor = conn.cursor()

# Create Leads Table
cursor.execute("""
CREATE TABLE IF NOT EXISTS leads (

    id INTEGER PRIMARY KEY AUTOINCREMENT,

    name TEXT,
    mobile TEXT,
    email TEXT,

    portfolio REAL,

    age INTEGER,
    income REAL,
    sip REAL,

    equity REAL,
    debt REAL,
    gold REAL,
    cash REAL

)
""")

# Save Changes
conn.commit()

# Close Database
conn.close()

print("Database Created Successfully")