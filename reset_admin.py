"""
Run this script once to reset the admin password.
Usage: python reset_admin.py
"""
import os
import urllib.parse
import pyodbc
from passlib.context import CryptContext
from dotenv import load_dotenv

load_dotenv(override=True)

DB_SERVER   = os.getenv("DB_SERVER")
DB_NAME     = os.getenv("DB_NAME")
DB_USERNAME = os.getenv("DB_USERNAME")
DB_PASSWORD = os.getenv("DB_PASSWORD")

NEW_PASSWORD = "Admin@123"   # change this if you want a different password
ADMIN_USER   = "admin"

conn_str = (
    f"DRIVER={{ODBC Driver 18 for SQL Server}};"
    f"SERVER={DB_SERVER};"
    f"DATABASE={DB_NAME};"
    f"UID={DB_USERNAME};"
    f"PWD={DB_PASSWORD};"
    f"TrustServerCertificate=yes;"
    f"Encrypt=no;"
    f"Connection Timeout=10;"
)

ctx = CryptContext(schemes=["bcrypt"], deprecated="auto")

try:
    conn   = pyodbc.connect(conn_str)
    cursor = conn.cursor()

    # Check if admin exists
    cursor.execute("SELECT Username, IsActive FROM SmsMonitorUsers_Vtb WHERE Username = ?", ADMIN_USER)
    row = cursor.fetchone()

    if not row:
        print(f"[!] User '{ADMIN_USER}' not found. Creating it now...")
        hashed = ctx.hash(NEW_PASSWORD)
        cursor.execute("""
            INSERT INTO SmsMonitorUsers_Vtb (Username, PasswordHash, Role, IsActive)
            VALUES (?, ?, 'admin', 1)
        """, ADMIN_USER, hashed)
        conn.commit()
        print(f"[OK] Admin user created. Login with: {ADMIN_USER} / {NEW_PASSWORD}")
    else:
        print(f"[INFO] Found user '{row[0]}', IsActive={bool(row[1])}")
        hashed = ctx.hash(NEW_PASSWORD)
        cursor.execute("""
            UPDATE SmsMonitorUsers_Vtb
            SET PasswordHash  = ?,
                IsActive      = 1,
                TokenVersion  = TokenVersion + 1
            WHERE Username = ?
        """, hashed, ADMIN_USER)
        conn.commit()
        print(f"[OK] Password reset. Login with: {ADMIN_USER} / {NEW_PASSWORD}")

    cursor.close()
    conn.close()

except Exception as e:
    print(f"[ERROR] {e}")
