import io
import os
import random
import smtplib
import sys
import threading

sys.setrecursionlimit(5000)
import urllib.parse
import pyodbc
from datetime import date, datetime, timedelta
from sqlalchemy import create_engine
from sqlalchemy.pool import QueuePool
from pathlib import Path
from dotenv import load_dotenv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.header import Header
from email import encoders
import email.utils
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

_BASE_DIR = Path(__file__).parent
load_dotenv(_BASE_DIR / ".env", override=True)
load_dotenv(_BASE_DIR / "email.env", override=True)

_last_synced_at = None   # updated each time sync_sms_usage() completes

REPORT_TEMPLATE_PATH = os.path.join(
    os.path.dirname(__file__), "report_templates", "SMS_Report_Template.xlsx"
)


def generate_report_from_template(client_name, kiosk_id, start_date, end_date, rows):
    """Fill the original XLSX template while preserving its legacy VML logo exactly."""
    import tempfile
    import zipfile
    import xml.etree.ElementTree as ET

    if not os.path.exists(REPORT_TEMPLATE_PATH):
        raise FileNotFoundError(
            f"Report template not found at {REPORT_TEMPLATE_PATH}. "
            "Place SMS_Report_Template.xlsx in the report_templates folder."
        )

    # openpyxl is used only to change workbook cell values.
    # The original template package parts that contain the logo are restored
    # afterward because this particular template stores the logo as a legacy
    # VML picture, not as a normal openpyxl image.
    wb = openpyxl.load_workbook(REPORT_TEMPLATE_PATH)
    ws = wb.active

    ws["A3"] = f"Client Name : {client_name}"
    ws["A4"] = f"ID : {kiosk_id}"
    ws["A5"] = f"From Date : {start_date}"
    ws["A6"] = f"To Date : {end_date}"

    start_row = 9
    row_count = 0
    for i, (sender_comp_id, mobile_no, message, sent_time) in enumerate(rows):
        r = start_row + i
        ws.cell(row=r, column=1, value=sender_comp_id)
        ws.cell(row=r, column=2, value=mobile_no)
        ws.cell(row=r, column=3, value=message)
        ws.cell(row=r, column=4, value=sent_time)
        ws.cell(row=r, column=5, value=len(message or ""))
        ws.cell(row=r, column=6, value=1)
        row_count += 1

    last_row = start_row + max(row_count, 1) - 1
    ws["F7"] = f"=SUM(F9:F{last_row})"

    fd, generated_path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)

    try:
        wb.save(generated_path)

        sheet_name = "xl/worksheets/sheet1.xml"
        sheet_rels_name = "xl/worksheets/_rels/sheet1.xml.rels"
        ct_name = "[Content_Types].xml"

        with zipfile.ZipFile(REPORT_TEMPLATE_PATH, "r") as template_zip, \
             zipfile.ZipFile(generated_path, "r") as generated_zip:

            template_names = set(template_zip.namelist())
            generated_names = set(generated_zip.namelist())
            replacements = {}

            # -------------------------------------------------------------
            # 1. Restore EVERY original drawing/VML/media package component.
            #    This template's visible logo is stored in:
            #      xl/drawings/vmlDrawing1.vml
            #      xl/drawings/_rels/vmlDrawing1.vml.rels
            #      xl/media/image1.png
            # -------------------------------------------------------------
            for name in template_names:
                if name.startswith("xl/drawings/") or name.startswith("xl/media/"):
                    replacements[name] = template_zip.read(name)

            # -------------------------------------------------------------
            # 2. Restore the original worksheet relationships.
            #    This contains:
            #      rId1 -> drawing1.xml
            #      rId2 -> vmlDrawing1.vml
            # -------------------------------------------------------------
            if sheet_rels_name in template_names:
                replacements[sheet_rels_name] = template_zip.read(sheet_rels_name)

            # -------------------------------------------------------------
            # 3. Restore BOTH worksheet references:
            #      <drawing r:id="rId1"/>
            #      <legacyDrawing r:id="rId2"/>
            #
            #    The second one is the critical reference for this template's
            #    logo. openpyxl removes it when saving the workbook.
            # -------------------------------------------------------------
            if sheet_name in template_names and sheet_name in generated_names:
                template_root = ET.fromstring(template_zip.read(sheet_name))
                generated_root = ET.fromstring(generated_zip.read(sheet_name))

                def local(tag):
                    return tag.rsplit("}", 1)[-1]

                # Find original drawing references.
                template_refs = {}
                for elem in template_root.iter():
                    name = local(elem.tag)
                    if name in {"drawing", "legacyDrawing"}:
                        template_refs[name] = ET.fromstring(
                            ET.tostring(elem, encoding="utf-8")
                        )

                # Remove any drawing/legacyDrawing nodes generated by openpyxl.
                for parent in generated_root.iter():
                    for child in list(parent):
                        if local(child.tag) in {"drawing", "legacyDrawing"}:
                            parent.remove(child)

                # Insert them at the proper worksheet positions.
                children = list(generated_root)

                # Standard drawing belongs near the end, before pageMargins.
                if "drawing" in template_refs:
                    insert_at = len(children)
                    for idx, child in enumerate(children):
                        if local(child.tag) == "pageMargins":
                            insert_at = idx
                            break
                    generated_root.insert(insert_at, template_refs["drawing"])

                # Legacy drawing normally belongs after drawing / near the
                # end of worksheet metadata, before pageMargins.
                if "legacyDrawing" in template_refs:
                    children = list(generated_root)
                    insert_at = len(children)
                    for idx, child in enumerate(children):
                        if local(child.tag) == "pageMargins":
                            insert_at = idx
                            break
                    generated_root.insert(insert_at, template_refs["legacyDrawing"])

                replacements[sheet_name] = ET.tostring(
                    generated_root,
                    encoding="utf-8",
                    xml_declaration=True
                )

            # -------------------------------------------------------------
            # 4. Restore required content-type declarations for VML/drawing.
            # -------------------------------------------------------------
            if ct_name in generated_names and ct_name in template_names:
                generated_ct = ET.fromstring(generated_zip.read(ct_name))
                template_ct = ET.fromstring(template_zip.read(ct_name))

                existing_parts = {
                    e.attrib.get("PartName") for e in generated_ct
                    if e.attrib.get("PartName")
                }
                existing_defaults = {
                    e.attrib.get("Extension") for e in generated_ct
                    if e.attrib.get("Extension")
                }

                for elem in template_ct:
                    part = elem.attrib.get("PartName", "")
                    ext = elem.attrib.get("Extension", "")

                    # Drawing override.
                    if part.startswith("/xl/drawings/") and part not in existing_parts:
                        generated_ct.append(
                            ET.fromstring(ET.tostring(elem, encoding="utf-8"))
                        )

                    # Image/VML defaults.
                    elif ext in {"vml", "png", "jpeg", "jpg", "gif", "bmp"} \
                            and ext not in existing_defaults:
                        generated_ct.append(
                            ET.fromstring(ET.tostring(elem, encoding="utf-8"))
                        )

                replacements[ct_name] = ET.tostring(
                    generated_ct,
                    encoding="utf-8",
                    xml_declaration=True
                )

            # -------------------------------------------------------------
            # 5. Build final workbook:
            #    generated workbook = data changes
            #    original template parts = logo and its relationships
            # -------------------------------------------------------------
            output = io.BytesIO()

            with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as final_zip:
                for item in generated_zip.infolist():
                    data = replacements.get(item.filename)
                    if data is None:
                        data = generated_zip.read(item.filename)
                    final_zip.writestr(item, data)

                # Add original package parts removed by openpyxl.
                for name, data in replacements.items():
                    if name not in generated_names:
                        final_zip.writestr(name, data)

            output.seek(0)
            return output, row_count

    finally:
        if os.path.exists(generated_path):
            os.remove(generated_path)

def generate_detailed_sms_report(client_id, start_date=None, end_date=None):
    """
    Generate the detailed, per-message Excel report (SMS_Report_Template.xlsx
    format) containing the actual SMS records from Smslog_Vtb for one client.
    Returns (excel_file, filename, sms_count) — or (None, None, 0) on failure.
    """
    conn = None
    cursor = None
    try:
        conn = get_connection()
        cursor = conn.cursor()

        cursor.execute("SELECT ClientName FROM ClientMst_Vtb WHERE ClientId = ?", client_id)
        name_row = cursor.fetchone()
        client_name = name_row[0].strip() if name_row else client_id

        query = """
            SELECT senderCompId, MobileNo, Message, SentTime
            FROM Smslog_Vtb
            WHERE senderCompId = ?
        """
        params = [client_id]
        if start_date:
            query += " AND CAST(SentTime AS DATE) >= ?"
            params.append(start_date)
        if end_date:
            query += " AND CAST(SentTime AS DATE) <= ?"
            params.append(end_date)
        query += " ORDER BY SentTime DESC"

        cursor.execute(query, params)
        rows = cursor.fetchall()
        cursor.close()
        conn.close()

        output, sms_count = generate_report_from_template(
            client_name=client_name,
            kiosk_id=client_id,
            start_date=start_date or "",
            end_date=end_date or "",
            rows=rows,
        )

        today_str = datetime.now().strftime("%Y%m%d")
        filename = f"sms_detailed_report_{client_id}_{today_str}.xlsx"

        print(f"[REPORT] Detailed SMS report generated: {client_id} - {sms_count} SMS")
        return output, filename, sms_count

    except Exception as e:
        print(f"[REPORT ERROR] generate_detailed_sms_report: {e}")
        try:
            if cursor: cursor.close()
            if conn: conn.close()
        except Exception:
            pass
        return None, None, 0

# ── In-memory TTL cache (60 s) — avoids hitting DB on every page load ───────
_CACHE_TTL = 60   # seconds

_dcache_clients:    dict = {"data": None, "ts": None}
_dcache_alerts:     dict = {"data": None, "ts": None}
_dcache_analytics:  dict = {}   # {client_id: {"data": ..., "ts": datetime}}
_dcache_thresholds: dict = {"data": None, "ts": None}

def _cache_fresh(entry: dict) -> bool:
    return (entry["ts"] is not None and
            (datetime.now() - entry["ts"]).total_seconds() < _CACHE_TTL)

def invalidate_dashboard_cache():
    """Call after any write that changes dashboard-visible data."""
    _dcache_clients["data"] = None
    _dcache_clients["ts"]   = None
    _dcache_alerts["data"]  = None
    _dcache_alerts["ts"]    = None
    _dcache_analytics.clear()


def get_thresholds() -> dict:
    """Return current warning/critical thresholds from DB (60 s cache).
    Falls back to hardcoded defaults (50/90) if the settings table has no rows yet."""
    if _cache_fresh(_dcache_thresholds) and _dcache_thresholds["data"] is not None:
        return _dcache_thresholds["data"]
    try:
        conn   = get_connection()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT SettingKey, SettingValue
            FROM SmsMonitorSettings_Vtb
            WHERE SettingKey IN ('threshold_warning', 'threshold_critical')
        """)
        rows   = {r[0]: int(r[1]) for r in cursor.fetchall()}
        cursor.close()
        conn.close()
        result = {
            "warning":  rows.get("threshold_warning",  50),
            "critical": rows.get("threshold_critical", 90),
        }
    except Exception as e:
        print(f"[DB ERROR] get_thresholds: {e}")
        result = {"warning": 50, "critical": 90}
    _dcache_thresholds["data"] = result
    _dcache_thresholds["ts"]   = datetime.now()
    return result


def set_thresholds(warning: int, critical: int, updated_by: str = None) -> bool:
    """Persist threshold values to DB and immediately invalidate the cache."""
    try:
        conn   = get_connection()
        cursor = conn.cursor()
        for key, val in [("threshold_warning", warning), ("threshold_critical", critical)]:
            cursor.execute("""
                MERGE SmsMonitorSettings_Vtb AS t
                USING (SELECT ? AS k, ? AS v, ? AS u) AS s(k, v, u)
                ON t.SettingKey = s.k
                WHEN MATCHED     THEN UPDATE SET SettingValue = s.v,
                                                 UpdatedBy    = s.u,
                                                 UpdatedDate  = GETDATE()
                WHEN NOT MATCHED THEN INSERT (SettingKey, SettingValue, UpdatedBy)
                                      VALUES (s.k, s.v, s.u);
            """, key, str(val), updated_by)
        conn.commit()
        cursor.close()
        conn.close()
        _dcache_thresholds["data"] = None
        _dcache_thresholds["ts"]   = None
        return True
    except Exception as e:
        print(f"[DB ERROR] set_thresholds: {e}")
        return False


def get_last_synced():
    """
    Returns ISO timestamp of the most recent sync.
    Uses the in-memory value when available (set after each sync_sms_usage run).
    Falls back to MAX(ModifiedDate) from ClientSmsCredit_Vtb so the footer
    always shows a real time instead of 'Not yet synced' after a server restart.
    """
    if _last_synced_at:
        return _last_synced_at.isoformat()
    try:
        conn   = get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "SELECT MAX(ModifiedDate) FROM ClientSmsCredit_Vtb WITH (NOLOCK)"
        )
        row = cursor.fetchone()
        cursor.close()
        conn.close()
        if row and row[0]:
            return row[0].isoformat()
    except Exception:
        pass
    return None

DB_SERVER   = os.getenv("DB_SERVER")
DB_NAME     = os.getenv("DB_NAME")
DB_USERNAME = os.getenv("DB_USERNAME")
DB_PASSWORD = os.getenv("DB_PASSWORD")

# ── Connection pool (SQLAlchemy QueuePool over pyodbc) ───────────────────────
# Built once at import time. get_connection() checks out a raw pyodbc connection
# from the pool; calling conn.close() returns it to the pool (not the DB).
#
#  pool_size    : idle connections kept open at all times
#  max_overflow : extra connections allowed under peak load (pool_size + max_overflow = hard cap)
#  pool_timeout : seconds to wait for a free connection before raising
#  pool_recycle : replace connections older than this (prevents SQL Server idle timeout drops)
#  pool_pre_ping: issues SELECT 1 before handing out a connection — silently replaces dead ones
_db_encrypt = os.getenv("DB_ENCRYPT", "no")   # set DB_ENCRYPT=yes in .env for encrypted SQL Server connections
_odbc_str = (
    f"DRIVER={{ODBC Driver 18 for SQL Server}};"
    f"SERVER={DB_SERVER};"
    f"DATABASE={DB_NAME};"
    f"UID={DB_USERNAME};"
    f"PWD={DB_PASSWORD};"
    f"TrustServerCertificate=yes;"
    f"Encrypt={_db_encrypt};"
    f"Connection Timeout=10;"
)
_engine = create_engine(
    "mssql+pyodbc:///?odbc_connect=" + urllib.parse.quote_plus(_odbc_str),
    poolclass    = QueuePool,
    pool_size    = 5,
    max_overflow = 10,
    pool_timeout = 30,
    pool_recycle = 1800,
    pool_pre_ping= True,
)


def get_connection():
    """Return a raw pyodbc connection checked out from the shared pool."""
    return _engine.raw_connection()


# ── Table reference ──────────────────────────────────────────
# ClientSmsUsageDaily_Vtb: RowId, ClientId (nvarchar), UsageDate, SmsCount, CreatedDate


def ensure_schema():
    """Creates missing tables and adds missing columns on startup."""
    try:
        conn   = get_connection()
        cursor = conn.cursor()

        # ── ClientSmsUsageDaily_Vtb ───────────────────────────────────
        cursor.execute("""
            IF NOT EXISTS (
                SELECT 1 FROM INFORMATION_SCHEMA.TABLES
                WHERE TABLE_NAME = 'ClientSmsUsageDaily_Vtb'
            )
            CREATE TABLE ClientSmsUsageDaily_Vtb (
                RowId       INT IDENTITY(1,1) PRIMARY KEY,
                ClientId    NVARCHAR(50)  NOT NULL,
                UsageDate   DATE          NOT NULL,
                SmsCount    INT           NOT NULL DEFAULT 0,
                CreatedDate DATETIME      NOT NULL DEFAULT GETDATE(),
                CONSTRAINT UQ_ClientDate UNIQUE (ClientId, UsageDate)
            )
        """)

        # ── SmsMonitorUsers_Vtb ───────────────────────────────────────
        cursor.execute("""
            IF NOT EXISTS (
                SELECT 1 FROM INFORMATION_SCHEMA.TABLES
                WHERE TABLE_NAME = 'SmsMonitorUsers_Vtb'
            )
            CREATE TABLE SmsMonitorUsers_Vtb (
                UserId       INT IDENTITY(1,1) PRIMARY KEY,
                Username     NVARCHAR(50)  NOT NULL UNIQUE,
                PasswordHash NVARCHAR(256) NOT NULL,
                Role         NVARCHAR(20)  NOT NULL DEFAULT 'viewer',
                IsActive     BIT           NOT NULL DEFAULT 1,
                CreatedDate  DATETIME      NOT NULL DEFAULT GETDATE(),
                LastLogin    DATETIME      NULL
            )
        """)

        # ── AuditLog_Vtb ──────────────────────────────────────────────
        cursor.execute("""
            IF NOT EXISTS (
                SELECT 1 FROM INFORMATION_SCHEMA.TABLES
                WHERE TABLE_NAME = 'AuditLog_Vtb'
            )
            CREATE TABLE AuditLog_Vtb (
                LogId       INT IDENTITY(1,1) PRIMARY KEY,
                Username    NVARCHAR(50)  NOT NULL,
                Action      NVARCHAR(100) NOT NULL,
                TargetId    NVARCHAR(50)  NULL,
                Detail      NVARCHAR(MAX) NULL,
                IpAddress   NVARCHAR(45)  NULL,
                CreatedDate DATETIME      NOT NULL DEFAULT GETDATE()
            )
        """)

        # ── AuditLog_Vtb indexes (idempotent) ─────────────────────────
        cursor.execute("""
            IF NOT EXISTS (
                SELECT 1 FROM sys.indexes
                WHERE object_id = OBJECT_ID('AuditLog_Vtb')
                  AND name = 'IX_AuditLog_CreatedDate'
            )
            CREATE INDEX IX_AuditLog_CreatedDate
                ON AuditLog_Vtb (CreatedDate DESC)
        """)
        cursor.execute("""
            IF NOT EXISTS (
                SELECT 1 FROM sys.indexes
                WHERE object_id = OBJECT_ID('AuditLog_Vtb')
                  AND name = 'IX_AuditLog_Username'
            )
            CREATE INDEX IX_AuditLog_Username
                ON AuditLog_Vtb (Username, CreatedDate DESC)
        """)

        # ── Seed default admin if no users exist ──────────────────────
        cursor.execute("SELECT COUNT(*) FROM SmsMonitorUsers_Vtb")
        user_count = cursor.fetchone()[0]
        if user_count == 0:
            from passlib.context import CryptContext
            _ctx         = CryptContext(schemes=["bcrypt"], deprecated="auto")
            default_user = os.getenv("DEFAULT_ADMIN_USER",     "admin")
            default_pass = os.getenv("DEFAULT_ADMIN_PASSWORD", "Admin@123")
            hashed       = _ctx.hash(default_pass)
            cursor.execute("""
                INSERT INTO SmsMonitorUsers_Vtb (Username, PasswordHash, Role)
                VALUES (?, ?, 'admin')
            """, default_user, hashed)
            print(f"[SCHEMA] Default admin user '{default_user}' created.")
            print(f"[SCHEMA] ⚠️  Set DEFAULT_ADMIN_PASSWORD in .env before going live!")

        # ── AlertSpikeSent / AlertDropSent columns ────────────────────
        cursor.execute("""
            IF NOT EXISTS (
                SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
                WHERE TABLE_NAME = 'ClientSmsCredit_Vtb'
                  AND COLUMN_NAME = 'AlertSpikeSent'
            )
            ALTER TABLE ClientSmsCredit_Vtb ADD AlertSpikeSent BIT NOT NULL DEFAULT 0
        """)
        cursor.execute("""
            IF NOT EXISTS (
                SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
                WHERE TABLE_NAME = 'ClientSmsCredit_Vtb'
                  AND COLUMN_NAME = 'AlertDropSent'
            )
            ALTER TABLE ClientSmsCredit_Vtb ADD AlertDropSent BIT NOT NULL DEFAULT 0
        """)

        # ── UsagePercent → computed persisted column ─────────────────
        # Converts UsagePercent from a manually-maintained stored column to a
        # SQL Server computed column so it can never drift out of sync with
        # UsedSMS / AllocatedSMS. Migration is idempotent: skipped if already
        # computed. Drops any default constraint before dropping the column.
        cursor.execute("""
            IF EXISTS (
                SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
                WHERE TABLE_NAME  = 'ClientSmsCredit_Vtb'
                  AND COLUMN_NAME = 'UsagePercent'
            )
            AND NOT EXISTS (
                SELECT 1 FROM sys.computed_columns
                WHERE object_id = OBJECT_ID('ClientSmsCredit_Vtb')
                  AND name      = 'UsagePercent'
            )
            BEGIN
                DECLARE @con NVARCHAR(256)
                SELECT @con = dc.name
                FROM   sys.default_constraints dc
                JOIN   sys.columns c
                       ON dc.parent_object_id = c.object_id
                      AND dc.parent_column_id = c.column_id
                WHERE  c.object_id = OBJECT_ID('ClientSmsCredit_Vtb')
                  AND  c.name = 'UsagePercent'
                IF @con IS NOT NULL
                    EXEC('ALTER TABLE ClientSmsCredit_Vtb DROP CONSTRAINT ' + @con)

                ALTER TABLE ClientSmsCredit_Vtb DROP COLUMN UsagePercent

                ALTER TABLE ClientSmsCredit_Vtb ADD UsagePercent AS
                    ROUND(
                        CASE WHEN AllocatedSMS > 0 THEN
                            CASE WHEN CAST(UsedSMS AS FLOAT) / AllocatedSMS * 100 > 999.99
                                 THEN 999.99
                                 ELSE CAST(UsedSMS AS FLOAT) / AllocatedSMS * 100
                            END
                        ELSE 0.0 END,
                    2) PERSISTED
            END
        """)

        # ── TokenVersion column ──────────────────────────────────────
        # Incremented on deactivate, password-reset, and logout so that
        # all existing JWTs for that user are immediately rejected.
        cursor.execute("""
            IF NOT EXISTS (
                SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
                WHERE TABLE_NAME = 'SmsMonitorUsers_Vtb'
                  AND COLUMN_NAME = 'TokenVersion'
            )
            ALTER TABLE SmsMonitorUsers_Vtb ADD TokenVersion INT NOT NULL DEFAULT 1
        """)

        # ── ManualAlloc column ────────────────────────────────────────
        # Marks rows where an admin explicitly set AllocatedSMS via the UI.
        # Sync never touches AllocatedSMS, only UsedSMS and RemainingSMS.
        cursor.execute("""
            IF NOT EXISTS (
                SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
                WHERE TABLE_NAME = 'ClientSmsCredit_Vtb'
                  AND COLUMN_NAME = 'ManualAlloc'
            )
            ALTER TABLE ClientSmsCredit_Vtb ADD ManualAlloc BIT NOT NULL DEFAULT 0
        """)

        # ── SmsMonitorSettings_Vtb ────────────────────────────────────
        cursor.execute("""
            IF NOT EXISTS (
                SELECT 1 FROM INFORMATION_SCHEMA.TABLES
                WHERE TABLE_NAME = 'SmsMonitorSettings_Vtb'
            )
            CREATE TABLE SmsMonitorSettings_Vtb (
                SettingKey   NVARCHAR(50)  NOT NULL PRIMARY KEY,
                SettingValue NVARCHAR(200) NOT NULL,
                UpdatedBy    NVARCHAR(50)  NULL,
                UpdatedDate  DATETIME      NOT NULL DEFAULT GETDATE()
            )
        """)

        # ── ValidityStartDate / ValidityEndDate columns ───────────────
        # Define the active window for each credit allocation. Usage and
        # threshold alerts are scoped to this window. Credits auto-reset
        # to 0 when ValidityEndDate passes (via expire_validity_credits).
        cursor.execute("""
            IF NOT EXISTS (
                SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
                WHERE TABLE_NAME = 'ClientSmsCredit_Vtb'
                  AND COLUMN_NAME = 'ValidityStartDate'
            )
            ALTER TABLE ClientSmsCredit_Vtb ADD ValidityStartDate DATE NULL
        """)
        cursor.execute("""
            IF NOT EXISTS (
                SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
                WHERE TABLE_NAME = 'ClientSmsCredit_Vtb'
                  AND COLUMN_NAME = 'ValidityEndDate'
            )
            ALTER TABLE ClientSmsCredit_Vtb ADD ValidityEndDate DATE NULL
        """)

        # ── ClientReportEmails_Vtb ────────────────────────────────
        cursor.execute("""
            IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='ClientReportEmails_Vtb')
            CREATE TABLE ClientReportEmails_Vtb (
                RowId       INT IDENTITY(1,1) PRIMARY KEY,
                ClientId    NVARCHAR(50)  NOT NULL UNIQUE,
                ClientName  NVARCHAR(200) NOT NULL,
                Email       NVARCHAR(1000) NOT NULL,
                CreatedDate DATETIME      NOT NULL DEFAULT GETDATE(),
                UpdatedDate DATETIME      NOT NULL DEFAULT GETDATE()
            )
        """)
        # Extend Email column if it was created with the old shorter length
        cursor.execute("""
            IF EXISTS (
                SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
                WHERE TABLE_NAME='ClientReportEmails_Vtb'
                  AND COLUMN_NAME='Email'
                  AND CHARACTER_MAXIMUM_LENGTH < 1000
            )
            ALTER TABLE ClientReportEmails_Vtb ALTER COLUMN Email NVARCHAR(1000) NOT NULL
        """)

        conn.commit()
        cursor.close()
        conn.close()
        print("[SCHEMA] Schema verified — all tables and columns OK.")
    except Exception as e:
        print(f"[SCHEMA ERROR] ensure_schema: {e}")


def get_user_by_username(username: str) -> dict | None:
    """Fetch a user row from SmsMonitorUsers_Vtb. Returns dict or None."""
    try:
        conn   = get_connection()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT Username, PasswordHash, Role, IsActive, TokenVersion
            FROM SmsMonitorUsers_Vtb
            WHERE Username = ?
        """, username)
        row = cursor.fetchone()
        cursor.close()
        conn.close()
        if not row:
            return None
        return {
            "username":      row[0],
            "password_hash": row[1],
            "role":          row[2],
            "is_active":     bool(row[3]),
            "token_version": int(row[4])
        }
    except Exception as e:
        print(f"[DB ERROR] get_user_by_username: {e}")
        return None


def update_last_login(username: str):
    """Stamp LastLogin on successful login."""
    try:
        conn   = get_connection()
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE SmsMonitorUsers_Vtb
            SET LastLogin = GETDATE()
            WHERE Username = ?
        """, username)
        conn.commit()
        cursor.close()
        conn.close()
    except Exception as e:
        print(f"[DB ERROR] update_last_login: {e}")


def log_audit(username: str, action: str, target_id: str = None,
              detail: str = None, ip_address: str = None):
    """Write one row to AuditLog_Vtb. Fire-and-forget — never raises."""
    try:
        conn   = get_connection()
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO AuditLog_Vtb (Username, Action, TargetId, Detail, IpAddress)
            VALUES (?, ?, ?, ?, ?)
        """, username, action, target_id, detail, ip_address)
        conn.commit()
        cursor.close()
        conn.close()
    except Exception as e:
        print(f"[AUDIT ERROR] {e}")


def fetch_clients():
    try:
        conn   = get_connection()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT ClientId, ClientName
            FROM ClientMst_Vtb
            WHERE ClientName IS NOT NULL
              AND LTRIM(RTRIM(ClientName)) != ''
            ORDER BY ClientName
        """)
        rows = cursor.fetchall()
        cursor.close()
        conn.close()
        return [{"id": row[0], "name": row[1].strip()} for row in rows]

    except pyodbc.Error as e:
        print(f"[DB ERROR] fetch_clients: {e}")
        return []


def assign_sms_credits(client_id, allocated_sms, start_date=None, end_date=None):
    """
    Assign SMS credits to a client.

    Validity-period mode (start_date + end_date provided):
      - REPLACES the existing allocation with a fresh start (no carryover).
      - Resets UsedSMS = 0 and clears all alert flags.
      - Usage tracking and threshold alerts are scoped to [start_date, end_date].

    Legacy cumulative mode (no dates):
      - Adds allocated_sms on top of any existing AllocatedSMS.

    Returns updated client record dict on success, None on failure.
    """
    try:
        conn   = get_connection()
        cursor = conn.cursor()

        # Check if client already has an allocation
        cursor.execute(
            "SELECT RowId, AllocatedSMS, UsedSMS FROM ClientSmsCredit_Vtb WHERE ClientId = ?",
            client_id
        )
        existing = cursor.fetchone()

        if existing:
            if start_date is not None:
                # Validity-period allocation: replace (fresh start, no carryover)
                new_allocated = allocated_sms
                used          = 0
                remaining     = new_allocated
                cursor.execute("""
                    UPDATE ClientSmsCredit_Vtb
                    SET AllocatedSMS      = ?,
                        UsedSMS           = 0,
                        RemainingSMS      = ?,
                        ManualAlloc       = 1,
                        Alert50Sent       = 0,
                        Alert90Sent       = 0,
                        AlertSpikeSent    = 0,
                        AlertDropSent     = 0,
                        ValidityStartDate = ?,
                        ValidityEndDate   = ?,
                        ModifiedDate      = GETDATE()
                    WHERE ClientId = ?
                """, new_allocated, remaining, start_date, end_date, client_id)
            else:
                # Legacy cumulative mode: add on top of existing
                new_allocated = existing[1] + allocated_sms
                used          = existing[2]
                remaining     = max(0, new_allocated - used)
                cursor.execute("""
                    UPDATE ClientSmsCredit_Vtb
                    SET AllocatedSMS = ?,
                        RemainingSMS = ?,
                        ManualAlloc  = 1,
                        ModifiedDate = GETDATE()
                    WHERE ClientId = ?
                """, new_allocated, remaining, client_id)

        else:
            new_allocated = allocated_sms
            used          = 0
            remaining     = new_allocated
            cursor.execute("""
                INSERT INTO ClientSmsCredit_Vtb
                    (ClientId, AllocatedSMS, UsedSMS, RemainingSMS,
                     Alert50Sent, Alert90Sent, ManualAlloc,
                     ValidityStartDate, ValidityEndDate,
                     CreatedDate, ModifiedDate)
                VALUES (?, ?, 0, ?, 0, 0, 1, ?, ?, GETDATE(), GETDATE())
            """, client_id, new_allocated, remaining, start_date, end_date)

        # Fetch client display name
        cursor.execute(
            "SELECT ClientName FROM ClientMst_Vtb WHERE ClientId = ?",
            client_id
        )
        name_row    = cursor.fetchone()
        client_name = name_row[0].strip() if name_row else client_id

        # Fetch today's SMS count
        cursor.execute("""
            SELECT ISNULL(SmsCount, 0) FROM ClientSmsUsageDaily_Vtb
            WHERE ClientId = ? AND UsageDate = CAST(GETDATE() AS DATE)
        """, client_id)
        today_row   = cursor.fetchone()
        today_usage = int(today_row[0]) if today_row else 0

        # Fetch yesterday's SMS count
        cursor.execute("""
            SELECT ISNULL(SmsCount, 0) FROM ClientSmsUsageDaily_Vtb
            WHERE ClientId = ? AND UsageDate = CAST(DATEADD(DAY, -1, GETDATE()) AS DATE)
        """, client_id)
        yest_row        = cursor.fetchone()
        yesterday_usage = int(yest_row[0]) if yest_row else 0

        # Fetch 7-day average (excluding today)
        cursor.execute("""
            SELECT ISNULL(AVG(CAST(SmsCount AS FLOAT)), 0)
            FROM ClientSmsUsageDaily_Vtb
            WHERE ClientId  = ?
              AND UsageDate >= CAST(DATEADD(DAY, -7, GETDATE()) AS DATE)
              AND UsageDate <  CAST(GETDATE() AS DATE)
        """, client_id)
        avg_row       = cursor.fetchone()
        seven_day_avg = float(avg_row[0]) if avg_row and avg_row[0] else 0.0

        spike_detected = (
            # Standard: today far exceeds recent average
            (seven_day_avg > 0 and today_usage > seven_day_avg * 2.0 and today_usage > yesterday_usage * 1.5)
            or
            # Re-activation: client was dormant, now sending real traffic
            (seven_day_avg <= 10 and yesterday_usage <= 10 and today_usage >= 50)
        )
        drop_detected = (
            yesterday_usage > 50 and
            seven_day_avg > 20 and
            today_usage < seven_day_avg * 0.5
        )

        conn.commit()
        cursor.close()
        conn.close()

        invalidate_dashboard_cache()   # new allocation changes dashboard data

        return {
            "id":               client_id,
            "name":             client_name,
            "allocated":        new_allocated,
            "used":             used,
            "remaining":        remaining,
            "usage_pct":        round(used / new_allocated * 100, 2) if new_allocated > 0 else 0.0,
            "today_usage":      today_usage,
            "yesterday_usage":  yesterday_usage,
            "spike_detected":   spike_detected,
            "drop_detected":    drop_detected,
            "validity_start":   start_date.isoformat() if start_date else None,
            "validity_end":     end_date.isoformat()   if end_date   else None,
        }

    except Exception as e:
        import traceback
        print("[ERROR] assign_sms_credits failed:")
        traceback.print_exc()
        return None


def fetch_dashboard_clients():
    """
    Returns all clients with assigned credits.
    Includes today/yesterday usage and spike/drop detection using Option 3 formula:

      Spike: TodayCount > 7DayAvg * 2.0  AND  TodayCount > YesterdayCount * 1.5
      Drop:  TodayCount < 7DayAvg * 0.5  AND  YesterdayCount > 50  AND  7DayAvg > 20

    Results are cached for 60 s — invalidated automatically after each sync.
    """
    if _cache_fresh(_dcache_clients) and _dcache_clients["data"] is not None:
        return _dcache_clients["data"]

    try:
        conn   = get_connection()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT
                c.ClientId,
                m.ClientName,
                c.AllocatedSMS,
                c.UsedSMS,
                c.RemainingSMS,
                c.UsagePercent,
                ISNULL(today.SmsCount, 0)   AS TodayUsage,
                ISNULL(yest.SmsCount,  0)   AS YesterdayUsage,
                ISNULL(avg7.AvgSms,    0)   AS SevenDayAvg,

                -- Spike: standard (today >> avg) OR re-activation (dormant client resumes)
                CASE
                    WHEN (
                        ISNULL(avg7.AvgSms, 0) > 0
                        AND ISNULL(today.SmsCount, 0) > avg7.AvgSms * 2.0
                        AND ISNULL(today.SmsCount, 0) > ISNULL(yest.SmsCount, 0) * 1.5
                    ) OR (
                        ISNULL(avg7.AvgSms,       0) <= 10
                        AND ISNULL(yest.SmsCount, 0) <= 10
                        AND ISNULL(today.SmsCount, 0) >= 50
                    )
                    THEN 1 ELSE 0
                END AS SpikeDetected,

                -- Drop: today < half of weekly avg AND yesterday was active AND avg was meaningful
                CASE
                    WHEN ISNULL(yest.SmsCount, 0) > 50
                     AND ISNULL(avg7.AvgSms,    0) > 20
                     AND ISNULL(today.SmsCount, 0) < avg7.AvgSms * 0.5
                    THEN 1 ELSE 0
                END AS DropDetected,

                c.ValidityStartDate,
                c.ValidityEndDate,
                re.Email AS ReportEmail

            FROM ClientSmsCredit_Vtb c WITH (NOLOCK)
            INNER JOIN ClientMst_Vtb m WITH (NOLOCK)
                ON c.ClientId = m.ClientId

            LEFT JOIN ClientSmsUsageDaily_Vtb today WITH (NOLOCK)
                ON today.ClientId  = c.ClientId
               AND today.UsageDate = CAST(GETDATE() AS DATE)

            LEFT JOIN ClientSmsUsageDaily_Vtb yest WITH (NOLOCK)
                ON yest.ClientId  = c.ClientId
               AND yest.UsageDate = CAST(DATEADD(DAY, -1, GETDATE()) AS DATE)

            LEFT JOIN (
                SELECT
                    ClientId,
                    AVG(CAST(SmsCount AS FLOAT)) AS AvgSms
                FROM ClientSmsUsageDaily_Vtb WITH (NOLOCK)
                WHERE UsageDate >= CAST(DATEADD(DAY, -7, GETDATE()) AS DATE)
                  AND UsageDate <  CAST(GETDATE() AS DATE)
                GROUP BY ClientId
            ) avg7 ON avg7.ClientId = c.ClientId

            LEFT JOIN ClientReportEmails_Vtb re WITH (NOLOCK)
                ON re.ClientId = c.ClientId

            ORDER BY m.ClientName
        """)
        rows = cursor.fetchall()
        cursor.close()
        conn.close()

        result = [
            {
                "id":               row[0],
                "name":             row[1].strip(),
                "allocated":        row[2],
                "used":             row[3],
                "remaining":        row[4],
                "usage_pct":        float(row[5]),
                "today_usage":      int(row[6]),
                "yesterday_usage":  int(row[7]),
                "seven_day_avg":    round(float(row[8]), 1),
                "spike_detected":   bool(row[9]),
                "drop_detected":    bool(row[10]),
                "validity_start":   row[11].strftime("%Y-%m-%d") if row[11] else None,
                "validity_end":     row[12].strftime("%Y-%m-%d") if row[12] else None,
                "report_email":     row[13],
            }
            for row in rows
        ]
        _dcache_clients["data"] = result
        _dcache_clients["ts"]   = datetime.now()
        return result

    except pyodbc.Error as e:
        print(f"[DB ERROR] fetch_dashboard_clients: {e}")
        return []


def fetch_active_alerts():
    """
    Returns all active alerts across clients:
      - threshold_90: UsagePercent >= 90
      - threshold_50: UsagePercent >= 50 and < 90
      - spike: today > 2x 7-day avg AND today > 1.5x yesterday
      - drop:  today < 0.5x 7-day avg AND yesterday > 50 AND avg > 20

    Results are cached for 60 s — invalidated automatically after each sync.
    """
    if _cache_fresh(_dcache_alerts) and _dcache_alerts["data"] is not None:
        return _dcache_alerts["data"]

    try:
        conn   = get_connection()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT
                c.ClientId,
                m.ClientName,
                c.UsagePercent,
                c.Alert50Sent,
                c.Alert90Sent,
                c.AlertSpikeSent,
                c.AlertDropSent,
                ISNULL(today.SmsCount, 0) AS TodayUsage,
                ISNULL(yest.SmsCount,  0) AS YestUsage,
                ISNULL(avg7.AvgSms,    0) AS SevenDayAvg,
                CASE
                    WHEN (
                        ISNULL(avg7.AvgSms, 0) > 0
                        AND ISNULL(today.SmsCount, 0) > avg7.AvgSms * 2.0
                        AND ISNULL(today.SmsCount, 0) > ISNULL(yest.SmsCount, 0) * 1.5
                    ) OR (
                        ISNULL(avg7.AvgSms,       0) <= 10
                        AND ISNULL(yest.SmsCount, 0) <= 10
                        AND ISNULL(today.SmsCount, 0) >= 50
                    )
                    THEN 1 ELSE 0
                END AS SpikeDetected,
                CASE
                    WHEN ISNULL(yest.SmsCount, 0) > 50
                     AND ISNULL(avg7.AvgSms,    0) > 20
                     AND ISNULL(today.SmsCount, 0) < avg7.AvgSms * 0.5
                    THEN 1 ELSE 0
                END AS DropDetected,
                c.ValidityStartDate,
                c.ValidityEndDate
            FROM ClientSmsCredit_Vtb c WITH (NOLOCK)
            INNER JOIN ClientMst_Vtb m WITH (NOLOCK) ON c.ClientId = m.ClientId
            LEFT JOIN ClientSmsUsageDaily_Vtb today WITH (NOLOCK)
                ON today.ClientId  = c.ClientId
               AND today.UsageDate = CAST(GETDATE() AS DATE)
            LEFT JOIN ClientSmsUsageDaily_Vtb yest WITH (NOLOCK)
                ON yest.ClientId  = c.ClientId
               AND yest.UsageDate = CAST(DATEADD(DAY, -1, GETDATE()) AS DATE)
            LEFT JOIN (
                SELECT ClientId, AVG(CAST(SmsCount AS FLOAT)) AS AvgSms
                FROM ClientSmsUsageDaily_Vtb WITH (NOLOCK)
                WHERE UsageDate >= CAST(DATEADD(DAY, -7, GETDATE()) AS DATE)
                  AND UsageDate <  CAST(GETDATE() AS DATE)
                GROUP BY ClientId
            ) avg7 ON avg7.ClientId = c.ClientId
            ORDER BY c.UsagePercent DESC, m.ClientName
        """)
        rows = cursor.fetchall()
        cursor.close()
        conn.close()

        t      = get_thresholds()
        alerts = []
        for row in rows:
            client_id        = row[0]
            client_name      = row[1].strip()
            usage_pct        = float(row[2])
            alert50_sent     = bool(row[3])
            alert90_sent     = bool(row[4])
            alert_spike_sent = bool(row[5])
            alert_drop_sent  = bool(row[6])
            today_usage      = int(row[7])
            yest_usage       = int(row[8])
            seven_day_avg    = round(float(row[9]), 1)
            spike            = bool(row[10])
            drop             = bool(row[11])
            validity_start   = row[12]
            validity_end     = row[13]

            today_d = date.today()
            within_validity = (
                validity_start is None or
                (validity_start <= today_d <= validity_end)
            )

            if usage_pct >= t["critical"] and within_validity:
                alerts.append({
                    "type":         "threshold_90",
                    "client_id":    client_id,
                    "client_name":  client_name,
                    "usage_pct":    round(usage_pct, 1),
                    "email_sent":   alert90_sent
                })
            elif usage_pct >= t["warning"] and within_validity:
                alerts.append({
                    "type":         "threshold_50",
                    "client_id":    client_id,
                    "client_name":  client_name,
                    "usage_pct":    round(usage_pct, 1),
                    "email_sent":   alert50_sent
                })

            if spike:
                alerts.append({
                    "type":           "spike",
                    "client_id":      client_id,
                    "client_name":    client_name,
                    "today_usage":    today_usage,
                    "seven_day_avg":  seven_day_avg,
                    "email_sent":     alert_spike_sent
                })

            if drop:
                alerts.append({
                    "type":           "drop",
                    "client_id":      client_id,
                    "client_name":    client_name,
                    "today_usage":    today_usage,
                    "yest_usage":     yest_usage,
                    "seven_day_avg":  seven_day_avg,
                    "email_sent":     alert_drop_sent
                })

        _dcache_alerts["data"] = alerts
        _dcache_alerts["ts"]   = datetime.now()
        return alerts

    except pyodbc.Error as e:
        print(f"[DB ERROR] fetch_active_alerts: {e}")
        return []


_MARK_ALERT_SQL = {
    "threshold_50": "UPDATE ClientSmsCredit_Vtb SET Alert50Sent    = 1 WHERE ClientId = ?",
    "threshold_90": "UPDATE ClientSmsCredit_Vtb SET Alert90Sent    = 1 WHERE ClientId = ?",
    "spike":        "UPDATE ClientSmsCredit_Vtb SET AlertSpikeSent = 1 WHERE ClientId = ?",
    "drop":         "UPDATE ClientSmsCredit_Vtb SET AlertDropSent  = 1 WHERE ClientId = ?",
}

def mark_alert_sent(client_id, alert_type):
    try:
        conn   = get_connection()
        cursor = conn.cursor()

        sql = _MARK_ALERT_SQL.get(alert_type)
        if sql:
            cursor.execute(sql, client_id)

        conn.commit()
        cursor.close()
        conn.close()
        return True

    except Exception as e:
        print(f"[DB ERROR] mark_alert_sent: {e}")
        return False


def expire_validity_credits():
    """
    Resets credits for all clients whose ValidityEndDate has passed.
    Called at midnight and on server startup. Leftover credits are NOT
    carried forward — the allocation is zeroed so a fresh assignment
    is needed for the next validity period.
    """
    try:
        conn   = get_connection()
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE ClientSmsCredit_Vtb
            SET AllocatedSMS      = 0,
                UsedSMS           = 0,
                RemainingSMS      = 0,
                Alert50Sent       = 0,
                Alert90Sent       = 0,
                AlertSpikeSent    = 0,
                AlertDropSent     = 0,
                ValidityStartDate = NULL,
                ValidityEndDate   = NULL,
                ModifiedDate      = GETDATE()
            WHERE ValidityEndDate IS NOT NULL
              AND CAST(GETDATE() AS DATE) > ValidityEndDate
              AND AllocatedSMS > 0
        """)
        expired = cursor.rowcount
        conn.commit()
        cursor.close()
        conn.close()
        if expired:
            print(f"[EXPIRE] Reset {expired} client(s) with expired validity periods.")
            invalidate_dashboard_cache()
        return expired
    except Exception as e:
        print(f"[DB ERROR] expire_validity_credits: {e}")
        return 0


def _split_emails(raw: str) -> list[str]:
    """Split a comma-separated email string into a clean list."""
    return [e.strip() for e in raw.split(",") if e.strip()]


def get_client_report_email(client_id: str) -> dict | None:
    """Return saved report emails for a client, or None."""
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "SELECT ClientId, ClientName, Email FROM ClientReportEmails_Vtb WHERE ClientId = ?",
            client_id
        )
        row = cursor.fetchone()
        cursor.close(); conn.close()
        if not row: return None
        return {"client_id": row[0], "client_name": row[1], "emails": _split_emails(row[2])}
    except Exception as e:
        print(f"[DB ERROR] get_client_report_email: {e}")
        return None


def save_client_report_email(client_id: str, client_name: str, emails: list[str]) -> bool:
    """Upsert the report emails (list) for a client, stored comma-separated."""
    email_str = ", ".join(e.strip() for e in emails if e.strip())
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("""
            MERGE ClientReportEmails_Vtb AS t
            USING (SELECT ? AS cid, ? AS cn, ? AS em) AS s(cid, cn, em)
            ON t.ClientId = s.cid
            WHEN MATCHED THEN
                UPDATE SET Email = s.em, ClientName = s.cn, UpdatedDate = GETDATE()
            WHEN NOT MATCHED THEN
                INSERT (ClientId, ClientName, Email, CreatedDate, UpdatedDate)
                VALUES (s.cid, s.cn, s.em, GETDATE(), GETDATE());
        """, client_id, client_name, email_str)
        conn.commit(); cursor.close(); conn.close()
        return True
    except Exception as e:
        print(f"[DB ERROR] save_client_report_email: {e}")
        return False


def fetch_all_client_report_emails() -> list:
    """Return all saved client report emails ordered by ClientName."""
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT ClientId, ClientName, Email, UpdatedDate
            FROM ClientReportEmails_Vtb
            ORDER BY ClientName
        """)
        rows = cursor.fetchall()
        cursor.close(); conn.close()
        return [
            {
                "client_id":   row[0],
                "client_name": row[1],
                "emails":      _split_emails(row[2]),
                "updated_at":  row[3].strftime("%d %b %Y") if row[3] else None,
            }
            for row in rows
        ]
    except Exception as e:
        print(f"[DB ERROR] fetch_all_client_report_emails: {e}")
        return []


def fetch_client_usage_report(client_id: str, start_date, end_date) -> dict | None:
    """Fetch day-wise SMS usage for a client between start_date and end_date (inclusive)."""
    try:
        conn = get_connection()
        cursor = conn.cursor()

        cursor.execute(
            "SELECT ClientName FROM ClientMst_Vtb WHERE ClientId = ?", client_id
        )
        name_row = cursor.fetchone()
        if not name_row:
            cursor.close(); conn.close()
            return None
        client_name = name_row[0].strip()

        cursor.execute("""
            SELECT UsageDate, ISNULL(SmsCount, 0)
            FROM ClientSmsUsageDaily_Vtb
            WHERE ClientId  = ?
              AND UsageDate >= ?
              AND UsageDate <= ?
            ORDER BY UsageDate
        """, client_id, start_date, end_date)
        rows = cursor.fetchall()
        cursor.close(); conn.close()

        daily = [
            {
                "date":  r[0].strftime("%Y-%m-%d"),
                "label": r[0].strftime("%d %b %Y"),
                "count": int(r[1])
            }
            for r in rows
        ]

        base = {
            "client_id":         client_id,
            "client_name":       client_name,
            "start_date_label":  start_date.strftime("%d %b %Y") if hasattr(start_date, "strftime") else str(start_date),
            "end_date_label":    end_date.strftime("%d %b %Y")   if hasattr(end_date,   "strftime") else str(end_date),
            "daily":             daily,
        }

        if not daily:
            return {**base, "total": 0, "avg": 0.0, "peak_day": None, "lowest_day": None}

        counts    = [d["count"] for d in daily]
        total     = sum(counts)
        avg       = round(total / len(counts), 1)
        peak_idx  = counts.index(max(counts))
        low_idx   = counts.index(min(counts))

        return {
            **base,
            "total":       total,
            "avg":         avg,
            "peak_day":    {"date": daily[peak_idx]["label"], "count": counts[peak_idx]},
            "lowest_day":  {"date": daily[low_idx]["label"],  "count": counts[low_idx]},
        }

    except Exception as e:
        print(f"[DB ERROR] fetch_client_usage_report: {e}")
        return None


def build_usage_report_excel(report: dict):
    """Build a professional Excel workbook for the usage report. Returns BytesIO."""
    import io
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SMS Usage Report"

    # ── Palette ──────────────────────────────────────────────────────
    NAVY   = "1E2A52"
    BLUE   = "30318C"
    TEAL   = "00B5A5"
    WHITE  = "FFFFFF"
    GREY1  = "F1F5F9"   # summary label bg / section header
    GREY2  = "F8FAFC"   # alt data row
    PEAK_C = "DBEAFE"   # peak row highlight
    LOW_C  = "DCFCE7"   # lowest row highlight
    TOT_C  = "E0E7FF"   # total row bg
    INK    = "1E293B"   # body text
    MUTED  = "64748B"   # secondary / row-number text
    LABEL  = "374151"   # summary label text

    def S(style, color):
        return Side(style=style, color=color)

    thin = S("thin",   "D1D5DB")
    teal_left = S("medium", TEAL)
    blue_med  = S("medium", BLUE)

    def bdr(l=None, r=None, t=None, b=None):
        return Border(
            left=l   or thin,
            right=r  or thin,
            top=t    or thin,
            bottom=b or thin,
        )

    B   = bdr()                                       # standard thin border
    BT  = bdr(t=blue_med, b=blue_med)                 # total row: bold top+bottom

    def fill(c):
        return PatternFill("solid", fgColor=c)

    def C(row, col, value="", bold=False, size=10, fg=INK, bg=None,
          ha="left", va="center", ind=0, border=None, fmt=None):
        c = ws.cell(row=row, column=col, value=value)
        c.font      = Font(name="Calibri", bold=bold, size=size, color=fg)
        if bg is not None:
            c.fill = fill(bg)
        c.alignment = Alignment(horizontal=ha, vertical=va, indent=ind)
        if border is not None:
            c.border = border
        if fmt is not None:
            c.number_format = fmt
        return c

    # ── Column layout ────────────────────────────────────────────────
    # A: row-number / accent bar (narrow)
    # B: label / date             (medium)
    # C: value / SMS count        (wide – must fit "01 Jun 2026  to  30 Jun 2026")
    # D: visual right-margin pad  (thin)
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 30
    ws.column_dimensions["D"].width = 3

    r = 1  # row cursor

    # ── Row 1: Brand bar (A:D merged) ────────────────────────────────
    ws.merge_cells(f"A{r}:D{r}")
    c1 = C(r, 1, "VAMS Global  ·  VAMS SMS Monitor",
           bold=True, size=11, fg=WHITE, bg=NAVY, ha="left", ind=1)
    c1.border = Border(left=S("thick", TEAL))
    ws.row_dimensions[r].height = 24
    r += 1

    # ── Row 2: Report title (A:D merged) ─────────────────────────────
    ws.merge_cells(f"A{r}:D{r}")
    C(r, 1, "Client SMS Usage Report",
      bold=True, size=15, fg=WHITE, bg=BLUE, ha="left", ind=1)
    ws.row_dimensions[r].height = 34
    r += 1

    # ── Row 3: thin teal accent stripe ───────────────────────────────
    ws.merge_cells(f"A{r}:D{r}")
    C(r, 1, bg=TEAL)
    ws.row_dimensions[r].height = 3
    r += 1

    # ── Row 4: spacer ────────────────────────────────────────────────
    ws.row_dimensions[r].height = 6
    r += 1

    # ── Row 5: "REPORT SUMMARY" section label ────────────────────────
    ws.merge_cells(f"A{r}:D{r}")
    C(r, 1, "  REPORT SUMMARY", bold=True, size=9,
      fg=MUTED, bg=GREY1, ha="left", va="center")
    ws.row_dimensions[r].height = 16
    r += 1

    # ── Summary rows: A=accent | B=label | C=value | D=margin ────────
    peak   = report.get("peak_day")
    lowest = report.get("lowest_day")

    summary = [
        ("Client Name",    report["client_name"],                                    False),
        ("Client ID",      report["client_id"],                                      False),
        ("Report Period",  f"{report['start_date_label']}  –  {report['end_date_label']}", False),
        ("Total SMS Sent", report["total"],                                           True),
        ("Daily Average",  round(report["avg"]),                                     True),
        ("Peak Day",       f"{peak['date']}   ({peak['count']:,} SMS)"     if peak   else "—", False),
        ("Lowest Day",     f"{lowest['date']}   ({lowest['count']:,} SMS)" if lowest else "—", False),
        ("Generated",      datetime.now().strftime("%d %b %Y,  %I:%M %p"),           False),
    ]

    for si, (lbl, val, is_num) in enumerate(summary):
        row_bg  = GREY1 if si % 2 == 0 else None
        # A: teal left-accent stripe
        acc = C(r, 1, bg=row_bg)
        acc.border = Border(left=teal_left, right=thin, top=thin, bottom=thin)
        # B: label
        C(r, 2, lbl, bold=True, size=10, fg=LABEL, bg=row_bg,
          ha="left", ind=1, border=bdr())
        # C: value
        C(r, 3, val, size=10, fg=INK, bg=row_bg,
          ha="right" if is_num else "left",
          ind=1, border=bdr(),
          fmt="#,##0" if is_num else None)
        # D: right margin
        C(r, 4, bg=row_bg, border=bdr())
        ws.row_dimensions[r].height = 20
        r += 1

    # ── Spacer ───────────────────────────────────────────────────────
    ws.row_dimensions[r].height = 10
    r += 1

    # ── Section label: "DAY-WISE BREAKDOWN" ──────────────────────────
    ws.merge_cells(f"A{r}:D{r}")
    C(r, 1, "  DAY-WISE SMS BREAKDOWN", bold=True, size=9,
      fg=MUTED, bg=GREY1, ha="left", va="center")
    ws.row_dimensions[r].height = 16
    r += 1

    # ── Table header ─────────────────────────────────────────────────
    HDR = r
    C(HDR, 1, "#",        bold=True, size=10, fg=WHITE, bg=BLUE,
      ha="center", va="center", border=bdr())
    C(HDR, 2, "Date",     bold=True, size=11, fg=WHITE, bg=BLUE,
      ha="left",   va="center", ind=1, border=bdr())
    C(HDR, 3, "SMS Used", bold=True, size=11, fg=WHITE, bg=BLUE,
      ha="right",  va="center", ind=1, border=bdr())
    C(HDR, 4, "",                            bg=BLUE, border=bdr())
    ws.row_dimensions[HDR].height = 26
    r += 1

    # ── Data rows ────────────────────────────────────────────────────
    peak_cnt   = peak["count"]   if peak   else None
    lowest_cnt = lowest["count"] if lowest else None

    for idx, day in enumerate(report["daily"]):
        count = day["count"]
        if   peak_cnt   is not None and count == peak_cnt:   bg = PEAK_C
        elif lowest_cnt is not None and count == lowest_cnt: bg = LOW_C
        elif idx % 2 == 1:                                   bg = GREY2
        else:                                                 bg = None

        C(r, 1, idx + 1, size=9, fg=MUTED, bg=GREY1,
          ha="center", va="center", border=bdr())
        C(r, 2, day["label"], size=10, fg=INK, bg=bg,
          ha="left", ind=1, border=bdr())
        C(r, 3, count, size=10, fg=INK, bg=bg,
          ha="right", ind=1, border=bdr(), fmt="#,##0")
        C(r, 4, "", bg=bg, border=bdr())
        ws.row_dimensions[r].height = 18
        r += 1

    # ── Total row ────────────────────────────────────────────────────
    C(r, 1, "",               bg=TOT_C, border=BT)
    C(r, 2, "TOTAL", bold=True, size=11, fg=BLUE, bg=TOT_C,
      ha="left", ind=1, border=BT)
    C(r, 3, report["total"], bold=True, size=11, fg=BLUE, bg=TOT_C,
      ha="right", ind=1, border=BT, fmt="#,##0")
    C(r, 4, "",               bg=TOT_C, border=BT)
    ws.row_dimensions[r].height = 24
    r += 1

    # ── Spacer ───────────────────────────────────────────────────────
    ws.row_dimensions[r].height = 8
    r += 1

    # ── Footer ───────────────────────────────────────────────────────
    ws.merge_cells(f"A{r}:D{r}")
    C(r, 1,
      f"Generated by VAMS SMS Monitor  ·  Confidential  ·  {datetime.now().strftime('%d %b %Y')}",
      size=8, fg=MUTED, bg=GREY1, ha="center", va="center")
    ws.row_dimensions[r].height = 16

    # ── Page setup (portrait, fit to one page wide) ──────────────────
    ws.page_setup.orientation  = "portrait"
    ws.page_setup.fitToPage    = True
    ws.page_setup.fitToWidth   = 1
    ws.page_setup.fitToHeight  = 0

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


class EmailNotifier:
    """Handles all alert email sending for SMS Monitor."""

    _COLORS = {
        "threshold_90": "#dc2626",
        "threshold_50": "#d97706",
        "spike":        "#7c3aed",
        "drop":         "#00B5A5",
    }
    _LABELS = {
        "threshold_90": "Critical Threshold Reached",
        "threshold_50": "Warning Threshold Reached",
        "spike":        "Sudden Usage Spike",
        "drop":         "Sudden Usage Drop",
    }
    _ICONS = {
        "threshold_90": "&#9888;",
        "threshold_50": "&#9888;",
        "spike":        "&#9650;",
        "drop":         "&#9660;",
    }

    def __init__(self):
        self.host      = os.getenv("SMTP_HOST",     "smtp.gmail.com")
        self.port      = int(os.getenv("SMTP_PORT",  587))
        self.user      = os.getenv("SMTP_USER",     "")
        self.password  = os.getenv("SMTP_PASSWORD", "")
        self.from_addr = os.getenv("ALERT_FROM", self.user)
        self.to_addrs  = [a.strip() for a in os.getenv("ALERT_TO", "").split(",") if a.strip()]
        self.cc_addrs  = [a.strip() for a in os.getenv("ALERT_CC", "").split(",") if a.strip()]
        self.timeout   = int(os.getenv("SMTP_TIMEOUT", 5))

    def _is_configured(self):
        return bool(self.user and self.password and self.to_addrs)

    _smtp_lock  = threading.Lock()
    _smtp_cache = {"server": None, "ts": None}
    _SMTP_TTL   = 270  # reuse connection for up to 4.5 min (Gmail idle timeout ~5 min)

    def _smtp_connect(self):
        if self.port == 465:
            server = smtplib.SMTP_SSL(self.host, self.port, timeout=self.timeout)
        else:
            server = smtplib.SMTP(self.host, self.port, timeout=self.timeout)
            server.ehlo()
            server.starttls()
            server.ehlo()
        server.login(self.user, self.password)
        return server

    def _get_smtp(self):
        """Returns a cached SMTP connection, reconnecting only when necessary."""
        with EmailNotifier._smtp_lock:
            srv = EmailNotifier._smtp_cache["server"]
            ts  = EmailNotifier._smtp_cache["ts"]
            if srv and ts and (datetime.now() - ts).total_seconds() < self._SMTP_TTL:
                try:
                    srv.noop()
                    EmailNotifier._smtp_cache["ts"] = datetime.now()
                    return srv
                except Exception:
                    pass
            try:
                if srv:
                    srv.quit()
            except Exception:
                pass
            new_srv = self._smtp_connect()
            EmailNotifier._smtp_cache["server"] = new_srv
            EmailNotifier._smtp_cache["ts"]     = datetime.now()
            return new_srv

    def _send(self, msg):
        """Send a message, retrying once with a fresh connection on disconnect."""
        try:
            self._get_smtp().send_message(msg)
        except smtplib.SMTPServerDisconnected:
            with EmailNotifier._smtp_lock:
                EmailNotifier._smtp_cache["server"] = None
                EmailNotifier._smtp_cache["ts"]     = None
            self._get_smtp().send_message(msg)

    def _table_row(self, label, value, value_color=None):
        vc = f"color:{value_color};" if value_color else ""
        return (
            f'<tr>'
            f'<td style="color:#666;padding:8px 0;border-bottom:1px solid #f0f0f0;width:140px">{label}</td>'
            f'<td style="font-weight:600;{vc}padding:8px 0;border-bottom:1px solid #f0f0f0">{value}</td>'
            f'</tr>'
        )

    def _build_html(self, alert_type, client_name, data):
        color = self._COLORS.get(alert_type, "#30318C")
        label = self._LABELS.get(alert_type, "Alert")
        icon  = self._ICONS.get(alert_type,  "!")
        tr    = self._table_row

        if alert_type in ("threshold_90", "threshold_50"):
            stats = (
                tr("Client",    client_name) +
                tr("Usage",     f"{data.get('usage_pct')}%",           color) +
                tr("Allocated", f"{data.get('allocated', 0):,} SMS") +
                tr("Used",      f"{data.get('used', 0):,} SMS") +
                tr("Remaining", f"{data.get('remaining', 0):,} SMS",   "#16a34a")
            )
        elif alert_type == "spike":
            avg   = data.get("seven_day_avg", 0)
            today = data.get("today_usage",   0)
            ratio = round(today / max(avg, 1), 1)
            stats = (
                tr("Client",        client_name) +
                tr("Today's Usage", f"{today:,} SMS",           color) +
                tr("7-Day Average", f"{avg:,} SMS") +
                tr("Spike Ratio",   f"{ratio}x above average",  color)
            )
        else:
            stats = (
                tr("Client",            client_name) +
                tr("Today's Usage",     f"{data.get('today_usage', 0):,} SMS",  color) +
                tr("Yesterday's Usage", f"{data.get('yest_usage', 0):,} SMS") +
                tr("7-Day Average",     f"{data.get('seven_day_avg', 0):,} SMS")
            )

        timestamp = datetime.now().strftime("%d %b %Y, %I:%M %p")
        return (
            f'<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f0f2f8;font-family:Arial,sans-serif;">'
            f'<div style="max-width:560px;margin:32px auto;background:#fff;border-radius:10px;'
            f'overflow:hidden;box-shadow:0 4px 20px rgba(35,41,96,0.12);">'

            # VAMS brand bar
            f'<div style="background:#232960;padding:13px 28px;display:flex;align-items:center;gap:10px;">'
            f'<div style="width:3px;height:28px;background:#00B5A5;border-radius:2px;flex-shrink:0;"></div>'
            f'<div>'
            f'<div style="font-size:10px;font-weight:700;color:#00B5A5;letter-spacing:2px;text-transform:uppercase;">VAMS Global</div>'
            f'<div style="font-size:13px;font-weight:700;color:#ffffff;letter-spacing:0.5px;">VAMS SMS Monitor</div>'
            f'</div>'
            f'</div>'

            # Alert type header
            f'<div style="background:{color};padding:20px 28px;">'
            f'<div style="font-size:20px;color:#fff;font-weight:700;margin-bottom:4px;">{icon}&nbsp;{label}</div>'
            f'<div style="color:rgba(255,255,255,0.85);font-size:13px;">Automated alert &middot; {client_name}</div>'
            f'</div>'

            # Body
            f'<div style="padding:24px 28px;">'
            f'<table style="width:100%;border-collapse:collapse;font-size:14px;">{stats}</table>'
            f'<div style="margin-top:20px;padding-top:12px;border-top:1px solid #f0f0f0;'
            f'color:#94a3b8;font-size:11px;">Detected at {timestamp}</div>'
            f'</div>'

            # VAMS footer
            f'<div style="background:#f8fafc;padding:12px 28px;border-top:1px solid #e8ecf4;">'
            f'<div style="font-size:11px;color:#64748b;">'
            f'<strong style="color:#232960;">VAMS Global</strong>'
            f'&nbsp;&middot;&nbsp;VAMS SMS Monitor'
            f'</div>'
            f'<div style="font-size:10px;color:#94a3b8;margin-top:3px;">This is an automated notification. Do not reply to this email.</div>'
            f'</div>'

            f'</div></body></html>'
        )

    def _stamp_headers(self, msg):
        """Add Date and Message-ID — missing these is a top spam trigger."""
        msg["Date"]       = email.utils.formatdate(localtime=True)
        msg["Message-ID"] = email.utils.make_msgid(domain="vamsglobal.com")

    _LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                              "static", "images", "email_logo.png")

    def _make_logo_part(self):
        """Return an inline MIMEImage for the VAMS logo, or None if unavailable."""
        try:
            with open(self._LOGO_PATH, "rb") as f:
                part = MIMEImage(f.read(), "png")
            part.add_header("Content-ID", "<vams_logo>")
            part.add_header("Content-Disposition", "inline", filename="vams_logo.png")
            return part
        except Exception:
            return None

    def _build_text(self, alert_type, client_name, data):
        label     = self._LABELS.get(alert_type, "Alert")
        timestamp = datetime.now().strftime("%d %b %Y, %I:%M %p")
        if alert_type in ("threshold_90", "threshold_50"):
            body = (
                f"Client    : {client_name}\n"
                f"Usage     : {data.get('usage_pct')}%\n"
                f"Allocated : {data.get('allocated', 0):,} SMS\n"
                f"Used      : {data.get('used', 0):,} SMS\n"
                f"Remaining : {data.get('remaining', 0):,} SMS"
            )
        elif alert_type == "spike":
            avg   = data.get("seven_day_avg", 0)
            today = data.get("today_usage", 0)
            body  = (
                f"Client        : {client_name}\n"
                f"Today's Usage : {today:,} SMS\n"
                f"7-Day Average : {avg:,} SMS\n"
                f"Spike Ratio   : {round(today / max(avg, 1), 1)}x above average"
            )
        else:
            body = (
                f"Client            : {client_name}\n"
                f"Today's Usage     : {data.get('today_usage', 0):,} SMS\n"
                f"Yesterday's Usage : {data.get('yest_usage', 0):,} SMS\n"
                f"7-Day Average     : {data.get('seven_day_avg', 0):,} SMS"
            )
        return (
            f"VAMS Global  |  VAMS SMS Monitor\n"
            f"{label}\n"
            f"{'─' * 40}\n"
            f"{body}\n"
            f"{'─' * 40}\n"
            f"Detected at {timestamp}\n"
            f"This is an automated notification. Do not reply."
        )

    def send(self, alert_type, client_name, data, _server=None):
        """Send a formatted HTML alert email. Returns True on success.
        Pass _server to reuse an already-open SMTP connection (avoids reconnect per email)."""
        if not self._is_configured():
            print("[EMAIL] SMTP credentials not configured in .env — skipping.")
            return False

        subjects = {
            "threshold_90": f"VAMS SMS Monitor: Critical Alert — {client_name} at {data.get('usage_pct')}%",
            "threshold_50": f"VAMS SMS Monitor: Warning — {client_name} at {data.get('usage_pct')}%",
            "spike":        f"VAMS SMS Monitor: Usage Spike Detected — {client_name}",
            "drop":         f"VAMS SMS Monitor: Usage Drop Detected — {client_name}",
        }

        msg = MIMEMultipart("alternative")
        msg["Subject"] = Header(subjects.get(alert_type, f"VAMS SMS Monitor: Alert — {client_name}"), "utf-8")
        msg["From"]    = self.from_addr
        msg["To"]      = ", ".join(self.to_addrs)
        if self.cc_addrs:
            msg["Cc"] = ", ".join(self.cc_addrs)
        self._stamp_headers(msg)
        msg.attach(MIMEText(self._build_text(alert_type, client_name, data), "plain", "utf-8"))
        msg.attach(MIMEText(self._build_html(alert_type, client_name, data), "html",  "utf-8"))

        all_recipients = self.to_addrs + self.cc_addrs
        try:
            if _server is not None:
                _server.send_message(msg)
            else:
                self._send(msg)
            print(f"[EMAIL] Sent '{alert_type}' for {client_name} to {', '.join(all_recipients)}")
            return True
        except Exception as e:
            print(f"[EMAIL ERROR] Failed '{alert_type}' for {client_name}: {e}")
            return False

    def dispatch_all(self):
        """Checks all clients for unsent alerts and sends them. Opens one SMTP connection
        for the entire batch so multiple alerts don't each pay a TCP+TLS handshake."""
        if not self._is_configured():
            return

        try:
            conn   = get_connection()
            cursor = conn.cursor()
            cursor.execute("""
                SELECT
                    c.ClientId,
                    m.ClientName,
                    c.UsagePercent,
                    c.AllocatedSMS,
                    c.UsedSMS,
                    c.RemainingSMS,
                    c.Alert50Sent,
                    c.Alert90Sent,
                    c.AlertSpikeSent,
                    c.AlertDropSent,
                    ISNULL(tod.SmsCount, 0),
                    ISNULL(yes.SmsCount, 0),
                    ISNULL(avg7.AvgSms,  0),
                    c.ValidityStartDate,
                    c.ValidityEndDate
                FROM ClientSmsCredit_Vtb c
                INNER JOIN ClientMst_Vtb m ON c.ClientId = m.ClientId
                LEFT JOIN ClientSmsUsageDaily_Vtb tod
                    ON tod.ClientId  = c.ClientId
                   AND tod.UsageDate = CAST(GETDATE() AS DATE)
                LEFT JOIN ClientSmsUsageDaily_Vtb yes
                    ON yes.ClientId  = c.ClientId
                   AND yes.UsageDate = CAST(DATEADD(DAY, -1, GETDATE()) AS DATE)
                LEFT JOIN (
                    SELECT ClientId, AVG(CAST(SmsCount AS FLOAT)) AS AvgSms
                    FROM ClientSmsUsageDaily_Vtb
                    WHERE UsageDate >= CAST(DATEADD(DAY, -7, GETDATE()) AS DATE)
                      AND UsageDate <  CAST(GETDATE() AS DATE)
                    GROUP BY ClientId
                ) avg7 ON avg7.ClientId = c.ClientId
            """)
            rows = cursor.fetchall()

            # Build the full list of emails to send before opening SMTP
            t       = get_thresholds()
            pending = []   # (alert_type, client_name, data, update_sql, client_id)
            for r in rows:
                cid        = r[0]
                name       = r[1].strip()
                usage_pct  = float(r[2])
                allocated  = int(r[3])
                used       = int(r[4])
                remaining  = int(r[5])
                a50_sent   = bool(r[6])
                a90_sent   = bool(r[7])
                spike_sent = bool(r[8])
                drop_sent  = bool(r[9])
                today          = int(r[10])
                yest           = int(r[11])
                avg            = round(float(r[12]), 1)
                validity_start = r[13]
                validity_end   = r[14]

                today_d = date.today()
                within_validity = (
                    validity_start is None or
                    (validity_start <= today_d <= validity_end)
                )

                spike = (
                    (avg > 0 and today > avg * 2.0 and today > yest * 1.5)
                    or (avg <= 10 and yest <= 10 and today >= 50)
                )
                drop = yest > 50 and avg > 20 and today < avg * 0.5

                if usage_pct >= t["critical"] and not a90_sent and within_validity:
                    pending.append(("threshold_90", name,
                                    {"usage_pct": round(usage_pct, 1), "allocated": allocated,
                                     "used": used, "remaining": remaining},
                                    "UPDATE ClientSmsCredit_Vtb SET Alert90Sent=1 WHERE ClientId=?", cid))
                elif usage_pct >= t["warning"] and not a50_sent and within_validity:
                    pending.append(("threshold_50", name,
                                    {"usage_pct": round(usage_pct, 1), "allocated": allocated,
                                     "used": used, "remaining": remaining},
                                    "UPDATE ClientSmsCredit_Vtb SET Alert50Sent=1 WHERE ClientId=?", cid))
                if spike and not spike_sent:
                    pending.append(("spike", name,
                                    {"today_usage": today, "seven_day_avg": avg},
                                    "UPDATE ClientSmsCredit_Vtb SET AlertSpikeSent=1 WHERE ClientId=?", cid))
                if drop and not drop_sent:
                    pending.append(("drop", name,
                                    {"today_usage": today, "yest_usage": yest, "seven_day_avg": avg},
                                    "UPDATE ClientSmsCredit_Vtb SET AlertDropSent=1 WHERE ClientId=?", cid))

            if pending:
                smtp = self._get_smtp()
                for alert_type, name, data, update_sql, cid in pending:
                    if self.send(alert_type, name, data, _server=smtp):
                            cursor.execute(update_sql, cid)

            conn.commit()
            cursor.close()
            conn.close()

        except Exception as e:
            print(f"[EMAIL ERROR] dispatch_all: {e}")

    def send_for_client(self, client_id, alert_type):
        """Manual send from dashboard button. Fetches client stats, sends, marks flag."""
        try:
            conn   = get_connection()
            cursor = conn.cursor()
            cursor.execute("""
                SELECT
                    m.ClientName,
                    c.UsagePercent,
                    c.AllocatedSMS,
                    c.UsedSMS,
                    c.RemainingSMS,
                    ISNULL(tod.SmsCount, 0),
                    ISNULL(yes.SmsCount, 0),
                    ISNULL(avg7.AvgSms,  0)
                FROM ClientSmsCredit_Vtb c WITH (NOLOCK)
                INNER JOIN ClientMst_Vtb m WITH (NOLOCK) ON c.ClientId = m.ClientId
                LEFT JOIN ClientSmsUsageDaily_Vtb tod WITH (NOLOCK)
                    ON tod.ClientId  = c.ClientId
                   AND tod.UsageDate = CAST(GETDATE() AS DATE)
                LEFT JOIN ClientSmsUsageDaily_Vtb yes WITH (NOLOCK)
                    ON yes.ClientId  = c.ClientId
                   AND yes.UsageDate = CAST(DATEADD(DAY, -1, GETDATE()) AS DATE)
                LEFT JOIN (
                    SELECT ClientId, AVG(CAST(SmsCount AS FLOAT)) AS AvgSms
                    FROM ClientSmsUsageDaily_Vtb WITH (NOLOCK)
                    WHERE UsageDate >= CAST(DATEADD(DAY, -7, GETDATE()) AS DATE)
                      AND UsageDate <  CAST(GETDATE() AS DATE)
                    GROUP BY ClientId
                ) avg7 ON avg7.ClientId = c.ClientId
                WHERE c.ClientId = ?
            """, client_id)
            row = cursor.fetchone()
            cursor.close()
            conn.close()

            if not row:
                return False

            name      = row[0].strip()
            usage_pct = round(float(row[1]), 1)
            allocated = int(row[2])
            used      = int(row[3])
            remaining = int(row[4])
            today     = int(row[5])
            yest      = int(row[6])
            avg       = round(float(row[7]), 1)

            data_map = {
                "threshold_90": {"usage_pct": usage_pct, "allocated": allocated,
                                 "used": used, "remaining": remaining},
                "threshold_50": {"usage_pct": usage_pct, "allocated": allocated,
                                 "used": used, "remaining": remaining},
                "spike":        {"today_usage": today, "seven_day_avg": avg},
                "drop":         {"today_usage": today, "yest_usage": yest,
                                 "seven_day_avg": avg},
            }

            sent = self.send(alert_type, name, data_map.get(alert_type, {}))
            if sent:
                mark_alert_sent(client_id, alert_type)
            return sent

        except Exception as e:
            print(f"[EMAIL ERROR] send_for_client: {e}")
            return False

    def send_client_report(self, client_id):
        """
        Manual per-client status report — works for ANY client regardless of alert status.
        Sends a full snapshot: usage %, allocated/used/remaining, today/yesterday/7-day avg,
        and active alert pills if any.
        """
        if not self._is_configured():
            print("[EMAIL] SMTP not configured — skipping client report.")
            return False

        try:
            conn   = get_connection()
            cursor = conn.cursor()
            cursor.execute("""
                SELECT
                    m.ClientName,
                    c.UsagePercent,
                    c.AllocatedSMS,
                    c.UsedSMS,
                    c.RemainingSMS,
                    ISNULL(tod.SmsCount, 0),
                    ISNULL(yes.SmsCount, 0),
                    ISNULL(avg7.AvgSms,  0)
                FROM ClientSmsCredit_Vtb c WITH (NOLOCK)
                INNER JOIN ClientMst_Vtb m WITH (NOLOCK) ON c.ClientId = m.ClientId
                LEFT JOIN ClientSmsUsageDaily_Vtb tod WITH (NOLOCK)
                    ON tod.ClientId  = c.ClientId
                   AND tod.UsageDate = CAST(GETDATE() AS DATE)
                LEFT JOIN ClientSmsUsageDaily_Vtb yes WITH (NOLOCK)
                    ON yes.ClientId  = c.ClientId
                   AND yes.UsageDate = CAST(DATEADD(DAY, -1, GETDATE()) AS DATE)
                LEFT JOIN (
                    SELECT ClientId, AVG(CAST(SmsCount AS FLOAT)) AS AvgSms
                    FROM ClientSmsUsageDaily_Vtb WITH (NOLOCK)
                    WHERE UsageDate >= CAST(DATEADD(DAY, -7, GETDATE()) AS DATE)
                      AND UsageDate <  CAST(GETDATE() AS DATE)
                    GROUP BY ClientId
                ) avg7 ON avg7.ClientId = c.ClientId
                WHERE c.ClientId = ?
            """, client_id)
            row = cursor.fetchone()
            cursor.close()
            conn.close()

            if not row:
                return False

            name      = row[0].strip()
            pct       = round(float(row[1]), 1)
            allocated = int(row[2])
            used      = int(row[3])
            remaining = int(row[4])
            today     = int(row[5])
            yest      = int(row[6])
            avg       = round(float(row[7]), 1)

            saved_email   = get_client_report_email(client_id)
            client_emails = saved_email["emails"] if saved_email else []

            spike = (
                (avg > 0 and today > avg * 2.0 and today > yest * 1.5)
                or (avg <= 10 and yest <= 10 and today >= 50)
            )
            drop = yest > 50 and avg > 20 and today < avg * 0.5

            t            = get_thresholds()
            status_color = "#dc2626" if pct >= t["critical"] else "#d97706" if pct >= t["warning"] else "#16a34a"
            status_label = "Critical"  if pct >= t["critical"] else "Warning"  if pct >= t["warning"] else "Healthy"
            header_color = "#dc2626" if pct >= t["critical"] else "#d97706" if pct >= t["warning"] else "#30318C"

            PILL_STYLE = {
                "critical": "background:#fef2f2;color:#dc2626;border:1px solid #fecaca",
                "warning":  "background:#fffbeb;color:#d97706;border:1px solid #fde68a",
                "spike":    "background:#f5f3ff;color:#7c3aed;border:1px solid #ddd6fe",
                "drop":     "background:#e6faf9;color:#00B5A5;border:1px solid #99e0da",
            }
            PILL_LABEL = {
                "critical": f"{t['critical']}% Threshold",
                "warning":  f"{t['warning']}% Threshold",
                "spike":    "&#9650; Spike",
                "drop":     "&#9660; Drop",
            }

            active_alerts = []
            if pct >= t["critical"]:   active_alerts.append("critical")
            elif pct >= t["warning"]:  active_alerts.append("warning")
            if spike:       active_alerts.append("spike")
            if drop:        active_alerts.append("drop")

            def pill_span(key):
                return (
                    f'<span style="display:inline-block;{PILL_STYLE[key]};padding:3px 10px;'
                    f'border-radius:100px;font-size:11px;font-weight:700;margin-right:6px;">'
                    f'{PILL_LABEL[key]}</span>'
                )

            pills_html = (
                "".join(pill_span(a) for a in active_alerts)
                if active_alerts else
                '<span style="color:#16a34a;font-weight:600;">&#10003; No active alerts</span>'
            )

            def tr_stat(label, value, color=None):
                vc = f"color:{color};font-weight:700;" if color else "color:#1e293b;"
                return (
                    f'<tr>'
                    f'<td style="color:#64748b;padding:8px 0;border-bottom:1px solid #f1f5f9;'
                    f'width:150px;font-size:13px;">{label}</td>'
                    f'<td style="padding:8px 0;border-bottom:1px solid #f1f5f9;font-size:13px;{vc}">{value}</td>'
                    f'</tr>'
                )

            stats_rows = (
                tr_stat("Status",        status_label,          status_color) +
                tr_stat("Usage",         f"{pct}%",             status_color) +
                tr_stat("Allocated",     f"{allocated:,} SMS") +
                tr_stat("Used",          f"{used:,} SMS") +
                tr_stat("Remaining",     f"{remaining:,} SMS",  "#16a34a") +
                tr_stat("Today",         f"{today:,} SMS") +
                tr_stat("Yesterday",     f"{yest:,} SMS") +
                tr_stat("7-Day Average", f"{avg:,} SMS")
            )

            timestamp   = datetime.now().strftime("%d %b %Y, %I:%M %p")
            report_date = datetime.now().strftime("%d %b %Y")

            html = (
                f'<!DOCTYPE html><html>'
                f'<body style="margin:0;padding:0;background:#f0f2f8;font-family:Arial,sans-serif;">'
                f'<div style="max-width:560px;margin:32px auto;background:#fff;border-radius:10px;'
                f'overflow:hidden;box-shadow:0 4px 20px rgba(35,41,96,0.12);">'

                # VAMS logo bar
                f'<div style="background:#ffffff;padding:18px 28px;border-bottom:3px solid #00B5A5;">'
                f'<img src="cid:vams_logo" alt="VAMS Global" style="height:46px;display:block;" /></div>'

                # Report header
                f'<div style="background:{header_color};padding:20px 28px;">'
                f'<div style="font-size:18px;color:#fff;font-weight:700;">Client Status Report</div>'
                f'<div style="color:rgba(255,255,255,0.82);font-size:13px;margin-top:4px;">'
                f'{name} &middot; {report_date}</div>'
                f'</div>'

                # Body
                f'<div style="padding:24px 28px;">'
                f'<div style="margin-bottom:16px;">{pills_html}</div>'
                f'<table style="width:100%;border-collapse:collapse;">{stats_rows}</table>'
                f'<div style="margin-top:20px;padding-top:12px;border-top:1px solid #f1f5f9;'
                f'color:#94a3b8;font-size:11px;">'
                f'Generated at {timestamp} &middot; Sent manually via VAMS SMS Monitor</div>'
                f'</div>'

                # VAMS footer
                f'<div style="background:#f8fafc;padding:12px 28px;border-top:1px solid #e8ecf4;">'
                f'<div style="font-size:11px;color:#64748b;">'
                f'<strong style="color:#232960;">VAMS Global</strong>'
                f'&nbsp;&middot;&nbsp;VAMS SMS Monitor'
                    f'</div>'
                f'<div style="font-size:10px;color:#94a3b8;margin-top:3px;">This is an automated notification. Do not reply to this email.</div>'
                f'</div>'

                f'</div></body></html>'
            )

            plain_text = (
                f"VAMS Global  |  VAMS SMS Monitor\n"
                f"Client Status Report — {name}  |  {report_date}\n"
                f"{'─' * 44}\n"
                f"Status        : {status_label}\n"
                f"Usage         : {pct}%\n"
                f"Allocated     : {allocated:,} SMS\n"
                f"Used          : {used:,} SMS\n"
                f"Remaining     : {remaining:,} SMS\n"
                f"Today         : {today:,} SMS\n"
                f"Yesterday     : {yest:,} SMS\n"
                f"7-Day Average : {avg:,} SMS\n"
                f"{'─' * 44}\n"
                f"Generated at {timestamp}\n"
                f"This is an automated notification. Do not reply."
            )
            if client_emails:
                to_addrs = list(client_emails)
                cc_addrs = [a for a in self.to_addrs if a not in client_emails]
            else:
                to_addrs = list(self.to_addrs)
                cc_addrs = []

            subject = f"VAMS SMS Monitor: Client Report — {name} ({report_date})"
            msg = MIMEMultipart("mixed")
            msg["Subject"] = Header(subject, "utf-8")
            msg["From"]    = self.from_addr
            msg["To"]      = ", ".join(to_addrs)
            if cc_addrs:
                msg["Cc"] = ", ".join(cc_addrs)
            self._stamp_headers(msg)

            related = MIMEMultipart("related")
            alt = MIMEMultipart("alternative")
            alt.attach(MIMEText(plain_text, "plain", "utf-8"))
            alt.attach(MIMEText(html,       "html",  "utf-8"))
            related.attach(alt)
            logo = self._make_logo_part()
            if logo:
                related.attach(logo)
            msg.attach(related)

            self._send(msg)
            cc_str = ", ".join(cc_addrs) if cc_addrs else "none"
            print(f"[EMAIL] Client report for {name} → To: {', '.join(to_addrs)} | CC: {cc_str}")
            return True

        except Exception as e:
            print(f"[EMAIL ERROR] send_client_report: {e}")
            return False

    def send_report_to_client_email(self, client_id: str) -> dict:
        """
        Send a detailed SMS usage report.
        Always sends to ALERT_TO (the system admin). If a client email is saved,
        it is added to the To line as well.
        REPORT_CC env var is added as CC when set.
        Returns {"ok": True} or {"ok": False, "reason": "smtp_error"|"no_data"}.
        """
        if not self._is_configured():
            return {"ok": False, "reason": "smtp_error"}

        # Client emails are optional — report goes to admin regardless
        saved         = get_client_report_email(client_id)
        client_emails = saved["emails"] if saved else []

        try:
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute("""
                SELECT
                    m.ClientName,
                    c.UsagePercent,
                    c.AllocatedSMS,
                    c.UsedSMS,
                    c.RemainingSMS,
                    c.Alert50Sent,
                    c.Alert90Sent,
                    c.ValidityStartDate,
                    c.ValidityEndDate,
                    c.ModifiedDate,
                    ISNULL(tod.SmsCount, 0),
                    ISNULL(yes.SmsCount, 0),
                    ISNULL(avg7.AvgSms,  0)
                FROM ClientSmsCredit_Vtb c WITH (NOLOCK)
                INNER JOIN ClientMst_Vtb m WITH (NOLOCK) ON c.ClientId = m.ClientId
                LEFT JOIN ClientSmsUsageDaily_Vtb tod WITH (NOLOCK)
                    ON tod.ClientId = c.ClientId
                   AND tod.UsageDate = CAST(GETDATE() AS DATE)
                LEFT JOIN ClientSmsUsageDaily_Vtb yes WITH (NOLOCK)
                    ON yes.ClientId = c.ClientId
                   AND yes.UsageDate = CAST(DATEADD(DAY,-1,GETDATE()) AS DATE)
                LEFT JOIN (
                    SELECT ClientId, AVG(CAST(SmsCount AS FLOAT)) AS AvgSms
                    FROM ClientSmsUsageDaily_Vtb WITH (NOLOCK)
                    WHERE UsageDate >= CAST(DATEADD(DAY,-7,GETDATE()) AS DATE)
                      AND UsageDate <  CAST(GETDATE() AS DATE)
                    GROUP BY ClientId
                ) avg7 ON avg7.ClientId = c.ClientId
                WHERE c.ClientId = ?
            """, client_id)
            row = cursor.fetchone()
            cursor.close(); conn.close()
        except Exception as e:
            print(f"[EMAIL ERROR] send_report_to_client_email DB: {e}")
            return {"ok": False, "reason": "smtp_error"}

        if not row:
            return {"ok": False, "reason": "no_data"}

        name      = row[0].strip()
        pct       = round(float(row[1]), 1)
        allocated = int(row[2])
        used      = int(row[3])
        remaining = int(row[4])
        a50       = bool(row[5])
        a90       = bool(row[6])
        v_start   = row[7]
        v_end     = row[8]
        modified  = row[9]
        today     = int(row[10])
        yest      = int(row[11])
        avg       = round(float(row[12]), 1)

        spike = (
            (avg > 0 and today > avg * 2.0 and today > yest * 1.5)
            or (avg <= 10 and yest <= 10 and today >= 50)
        )
        drop = yest > 50 and avg > 20 and today < avg * 0.5

        t            = get_thresholds()
        status_label = "Critical" if pct >= t["critical"] else "Warning" if pct >= t["warning"] else "Healthy"
        status_color = "#dc2626"  if pct >= t["critical"] else "#d97706" if pct >= t["warning"] else "#16a34a"
        header_color = "#dc2626"  if pct >= t["critical"] else "#d97706" if pct >= t["warning"] else "#30318C"

        def fd(d):
            if not d: return "—"
            return d.strftime("%d %b %Y") if hasattr(d, "strftime") else str(d)

        validity_str = f"{fd(v_start)} – {fd(v_end)}" if v_start else "No validity period set"
        last_updated = modified.strftime("%d %b %Y, %I:%M %p") if modified else "—"
        report_date  = datetime.now().strftime("%d %b %Y")
        timestamp    = datetime.now().strftime("%d %b %Y, %I:%M %p")

        def tr(label, value, color=None):
            vc = f"color:{color};font-weight:700;" if color else "color:#1e293b;"
            return (
                f'<tr><td style="color:#64748b;padding:8px 0;border-bottom:1px solid #f1f5f9;'
                f'width:180px;font-size:13px;">{label}</td>'
                f'<td style="padding:8px 0;border-bottom:1px solid #f1f5f9;font-size:13px;{vc}">{value}</td></tr>'
            )

        def pill(bg, color, border, text):
            return (f'<span style="display:inline-block;background:{bg};color:{color};'
                    f'border:1px solid {border};padding:3px 10px;border-radius:100px;'
                    f'font-size:11px;font-weight:700;margin:2px 4px 2px 0;">{text}</span>')

        # Alert pills
        pills = ""
        if pct >= t["critical"]:   pills += pill("#fef2f2","#dc2626","#fecaca", f"⚠ Critical — {pct}%")
        elif pct >= t["warning"]:  pills += pill("#fffbeb","#d97706","#fde68a", f"⚠ Warning — {pct}%")
        else:                      pills += pill("#f0fdf4","#16a34a","#bbf7d0", f"✓ Healthy — {pct}%")
        if a50: pills += pill("#fffbeb","#d97706","#fde68a","50% Threshold Crossed")
        if a90: pills += pill("#fef2f2","#dc2626","#fecaca","90% Threshold Crossed")
        if spike: pills += pill("#f5f3ff","#7c3aed","#ddd6fe","▲ Usage Spike")
        if drop:  pills += pill("#e6faf9","#00B5A5","#99e0da","▼ Usage Drop")
        if not (a50 or a90 or spike or drop) and pct < t["warning"]:
            pills += '<span style="font-size:12px;color:#16a34a;font-weight:600;">✓ No alerts — usage is within normal range</span>'

        # Threshold notes for plain text
        notes = []
        if pct >= t["critical"]:  notes.append(f"CRITICAL: {pct}% usage — {t['critical']}% threshold exceeded")
        elif pct >= t["warning"]: notes.append(f"WARNING: {pct}% usage — {t['warning']}% threshold exceeded")
        else:                     notes.append(f"Normal: {pct}% usage — no threshold crossed")
        if a50:   notes.append("50% threshold was previously crossed (alert sent)")
        if a90:   notes.append("90% threshold was previously crossed (alert sent)")
        if spike: notes.append(f"Usage spike detected: today ({today:,}) is {round(today/max(avg,1),1)}x the 7-day avg ({avg:,})")
        if drop:  notes.append(f"Usage drop detected: today ({today:,}) vs yesterday ({yest:,}) and avg ({avg:,})")

        stats = (
            tr("Status",        status_label,           status_color) +
            tr("Usage",         f"{pct}%",              status_color) +
            tr("Allocated",     f"{allocated:,} SMS") +
            tr("Used",          f"{used:,} SMS") +
            tr("Remaining",     f"{remaining:,} SMS",   "#16a34a") +
            tr("Today",         f"{today:,} SMS") +
            tr("Yesterday",     f"{yest:,} SMS") +
            tr("7-Day Average", f"{avg:,} SMS") +
            tr("Validity",      validity_str) +
            tr("Last Updated",  last_updated)
        )

        html = (
            f'<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f0f2f8;font-family:Arial,sans-serif;">'
            f'<div style="max-width:580px;margin:32px auto;background:#fff;border-radius:10px;'
            f'overflow:hidden;box-shadow:0 4px 20px rgba(35,41,96,0.12);">'
            f'<div style="background:#ffffff;padding:18px 28px;border-bottom:3px solid #00B5A5;">'
            f'<img src="cid:vams_logo" alt="VAMS Global" style="height:46px;display:block;" /></div>'
            f'<div style="background:{header_color};padding:20px 28px;">'
            f'<div style="font-size:18px;color:#fff;font-weight:700;">Client SMS Usage Report</div>'
            f'<div style="color:rgba(255,255,255,0.82);font-size:13px;margin-top:4px;">{name} &middot; {report_date}</div></div>'
            f'<div style="padding:16px 28px 4px;">{pills}</div>'
            f'<div style="padding:4px 28px 24px;">'
            f'<table style="width:100%;border-collapse:collapse;">{stats}</table>'
            f'<div style="margin-top:16px;padding-top:12px;border-top:1px solid #f1f5f9;color:#94a3b8;font-size:11px;">'
            f'Generated at {timestamp} &middot; Sent via VAMS SMS Monitor</div></div>'
            f'<div style="background:#f8fafc;padding:12px 28px;border-top:1px solid #e8ecf4;">'
            f'<div style="font-size:11px;color:#64748b;"><strong style="color:#232960;">VAMS Global</strong>'
            f'&nbsp;&middot;&nbsp;VAMS SMS Monitor</div>'
            f'<div style="font-size:10px;color:#94a3b8;margin-top:3px;">This is an automated notification. Do not reply to this email.</div>'
            f'</div></div></body></html>'
        )

        plain = (
            f"VAMS Global  |  VAMS SMS Monitor\n"
            f"Client SMS Usage Report — {name}  |  {report_date}\n"
            f"{'─'*44}\n"
            f"Status        : {status_label}\n"
            f"Usage         : {pct}%\n"
            f"Allocated     : {allocated:,} SMS\n"
            f"Used          : {used:,} SMS\n"
            f"Remaining     : {remaining:,} SMS\n"
            f"Today         : {today:,} SMS\n"
            f"Yesterday     : {yest:,} SMS\n"
            f"7-Day Average : {avg:,} SMS\n"
            f"Validity      : {validity_str}\n"
            f"Last Updated  : {last_updated}\n"
            f"{'─'*44}\n"
            f"Alerts:\n" + "\n".join(f"  • {n}" for n in notes) +
            f"\n{'─'*44}\n"
            f"Generated at {timestamp}\n"
            f"This is an automated notification. Do not reply."
        )

        if client_emails:
            to_addrs = list(client_emails)
            cc_addrs = [a for a in self.to_addrs if a not in to_addrs]
        else:
            to_addrs = list(self.to_addrs)
            cc_addrs = []

        subject = f"VAMS SMS Monitor: Client Report — {name} ({report_date})"
        msg = MIMEMultipart("mixed")
        msg["Subject"] = Header(subject, "utf-8")
        msg["From"]    = self.from_addr
        msg["To"]      = ", ".join(to_addrs)
        if cc_addrs:
            msg["Cc"] = ", ".join(cc_addrs)
        self._stamp_headers(msg)

        related = MIMEMultipart("related")
        alt = MIMEMultipart("alternative")
        alt.attach(MIMEText(plain, "plain", "utf-8"))
        alt.attach(MIMEText(html,  "html",  "utf-8"))
        related.attach(alt)
        logo = self._make_logo_part()
        if logo:
            related.attach(logo)
        msg.attach(related)

        try:
            self._send(msg)
            cc_str = ", ".join(cc_addrs) if cc_addrs else "none"
            print(f"[EMAIL] Client report for {name} → To: {', '.join(to_addrs)} | CC: {cc_str}")
            return {"ok": True}
        except Exception as e:
            print(f"[EMAIL ERROR] send_report_to_client_email: {e}")
            return {"ok": False, "reason": "smtp_error"}

    def send_usage_report_email(self, client_id: str, start_date, end_date) -> dict:
        """
        Send day-wise SMS usage report. If a client email is saved: To = client, CC = admin.
        If no client email is saved: falls back to admin address as To.
        Returns {"ok": True} or {"ok": False, "reason": "no_data"|"smtp_error"}.
        """
        if not self._is_configured():
            return {"ok": False, "reason": "smtp_error"}

        saved     = get_client_report_email(client_id)
        to_emails = saved["emails"] if saved else []

        report = fetch_client_usage_report(client_id, start_date, end_date)
        if not report or not report.get("daily"):
            return {"ok": False, "reason": "no_data"}

        name         = report["client_name"]
        total        = report["total"]
        avg          = report["avg"]
        peak         = report.get("peak_day")
        lowest       = report.get("lowest_day")
        daily        = report["daily"]
        start_lbl    = report["start_date_label"]
        end_lbl      = report["end_date_label"]
        timestamp    = datetime.now().strftime("%d %b %Y, %I:%M %p")

        # ── HTML table rows ────────────────────────────────────────────────
        def tr(label, val, color=None):
            vc = f"color:{color};font-weight:700;" if color else "color:#1e293b;"
            return (
                f'<tr><td style="color:#64748b;padding:7px 0;border-bottom:1px solid #f1f5f9;'
                f'width:160px;font-size:13px;">{label}</td>'
                f'<td style="padding:7px 0;border-bottom:1px solid #f1f5f9;font-size:13px;{vc}">{val}</td></tr>'
            )

        summary_rows = (
            tr("Client ID",     report["client_id"]) +
            tr("Client Name",   name) +
            tr("Period",        f"{start_lbl} – {end_lbl}") +
            tr("Total SMS",     f"{total:,}") +
            tr("Daily Average", f"{avg:,}") +
            tr("Peak Day",      f"{peak['date']} ({peak['count']:,} SMS)"     if peak   else "—", "#30318C") +
            tr("Lowest Day",    f"{lowest['date']} ({lowest['count']:,} SMS)" if lowest else "—")
        )

        # ── Day-wise table ─────────────────────────────────────────────────
        day_rows = "".join(
            f'<tr><td style="padding:6px 12px;border-bottom:1px solid #f8fafc;font-size:12px;">{d["label"]}</td>'
            f'<td style="padding:6px 12px;border-bottom:1px solid #f8fafc;font-size:12px;text-align:right;">{d["count"]:,}</td></tr>'
            for d in daily
        )

        html = (
            f'<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f0f2f8;font-family:Arial,sans-serif;">'
            f'<div style="max-width:600px;margin:32px auto;background:#fff;border-radius:10px;'
            f'overflow:hidden;box-shadow:0 4px 20px rgba(35,41,96,0.12);">'
            f'<div style="background:#ffffff;padding:18px 28px;border-bottom:3px solid #00B5A5;">'
            f'<img src="cid:vams_logo" alt="VAMS Global" style="height:46px;display:block;" /></div>'
            f'<div style="background:#30318C;padding:20px 28px;">'
            f'<div style="font-size:18px;color:#fff;font-weight:700;">Client SMS Usage Report</div>'
            f'<div style="color:rgba(255,255,255,0.8);font-size:13px;margin-top:4px;">{name} &middot; {start_lbl} – {end_lbl}</div></div>'
            f'<div style="padding:20px 28px;">'
            f'<table style="width:100%;border-collapse:collapse;">{summary_rows}</table>'
            f'<div style="margin:20px 0 10px;font-weight:700;font-size:13px;color:#0f172a;">Day-wise Usage</div>'
            f'<table style="width:100%;border-collapse:collapse;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden;">'
            f'<thead><tr>'
            f'<th style="background:#f8fafc;padding:8px 12px;text-align:left;font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:.5px;">Date</th>'
            f'<th style="background:#f8fafc;padding:8px 12px;text-align:right;font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:.5px;">SMS Used</th>'
            f'</tr></thead><tbody>{day_rows}</tbody></table>'
            f'<div style="margin-top:16px;color:#94a3b8;font-size:11px;">Generated at {timestamp}</div>'
            f'</div>'
            f'<div style="background:#f8fafc;padding:12px 28px;border-top:1px solid #e8ecf4;">'
            f'<div style="font-size:11px;color:#64748b;"><strong style="color:#232960;">VAMS Global</strong> &middot; VAMS SMS Monitor</div>'
            f'<div style="font-size:10px;color:#94a3b8;margin-top:3px;">This is an automated report. Do not reply to this email.</div>'
            f'</div></div></body></html>'
        )

        plain = (
            f"VAMS Global  |  VAMS SMS Monitor\n"
            f"Client SMS Usage Report — {name}  |  {start_lbl} – {end_lbl}\n"
            f"{'─'*48}\n"
            f"Client ID     : {report['client_id']}\n"
            f"Client Name   : {name}\n"
            f"Period        : {start_lbl} – {end_lbl}\n"
            f"Total SMS     : {total:,}\n"
            f"Daily Average : {avg:,}\n"
            f"Peak Day      : {peak['date']} ({peak['count']:,} SMS)" + (f"\n" if peak else "\n") +
            f"Lowest Day    : {lowest['date']} ({lowest['count']:,} SMS)" + (f"\n" if lowest else "\n") +
            f"{'─'*48}\n"
            f"Day-wise Usage:\n" +
            "\n".join(f"  {d['label']}: {d['count']:,} SMS" for d in daily) +
            f"\n{'─'*48}\n"
            f"Generated at {timestamp}\n"
            f"This is an automated report. Do not reply."
        )

        # Build Excel attachment
        buf = build_usage_report_excel(report)

        subject = f"VAMS SMS Monitor: Usage Report — {name} ({start_lbl} to {end_lbl})"
        msg = MIMEMultipart("mixed")
        msg["Subject"] = Header(subject, "utf-8")
        msg["From"]    = self.from_addr
        if to_emails:
            msg["To"] = ", ".join(to_emails)
            cc_addrs  = [a for a in self.to_addrs if a not in to_emails]
            if cc_addrs:
                msg["Cc"] = ", ".join(cc_addrs)
        else:
            msg["To"] = ", ".join(self.to_addrs)
            cc_addrs  = []
        self._stamp_headers(msg)

        related = MIMEMultipart("related")
        alt = MIMEMultipart("alternative")
        alt.attach(MIMEText(plain, "plain", "utf-8"))
        alt.attach(MIMEText(html,  "html",  "utf-8"))
        related.attach(alt)
        logo = self._make_logo_part()
        if logo:
            related.attach(logo)
        msg.attach(related)

        safe_name = name.replace(" ", "_")
        filename  = f"sms_usage_{safe_name}_{start_date}_{end_date}.xlsx"
        part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        part.set_payload(buf.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
        msg.attach(part)

        try:
            self._send(msg)
            recipients = to_emails if to_emails else list(self.to_addrs)
            print(f"[EMAIL] Usage report for {name} sent to {', '.join(recipients)}")
            return {"ok": True, "emails": recipients}
        except Exception as e:
            print(f"[EMAIL ERROR] send_usage_report_email: {e}")
            return {"ok": False, "reason": "smtp_error"}

    def send_detailed_usage_report_email(self, client_id: str, start_date, end_date) -> dict:
        """
        Same recipients / HTML summary body as send_usage_report_email, but the
        attachment is the DETAILED per-message report (SMS_Report_Template.xlsx
        format — senderCompId / MobileNo / Message / SentTime / SMS Lenth / SMSCount)
        instead of the day-wise summary workbook.
        Returns {"ok": True, "emails": [...], "sms_count": int} or
        {"ok": False, "reason": "no_data"|"smtp_error"}.
        """
        if not self._is_configured():
            return {"ok": False, "reason": "smtp_error"}

        saved     = get_client_report_email(client_id)
        to_emails = saved["emails"] if saved else []

        report = fetch_client_usage_report(client_id, start_date, end_date)
        if not report or not report.get("daily"):
            return {"ok": False, "reason": "no_data"}

        excel_buf, filename, sms_count = generate_detailed_sms_report(client_id, start_date, end_date)
        if excel_buf is None:
            return {"ok": False, "reason": "smtp_error"}

        name         = report["client_name"]
        total        = report["total"]
        avg          = report["avg"]
        peak         = report.get("peak_day")
        lowest       = report.get("lowest_day")
        daily        = report["daily"]
        start_lbl    = report["start_date_label"]
        end_lbl      = report["end_date_label"]
        timestamp    = datetime.now().strftime("%d %b %Y, %I:%M %p")

        def tr(label, val, color=None):
            vc = f"color:{color};font-weight:700;" if color else "color:#1e293b;"
            return (
                f'<tr><td style="color:#64748b;padding:7px 0;border-bottom:1px solid #f1f5f9;'
                f'width:160px;font-size:13px;">{label}</td>'
                f'<td style="padding:7px 0;border-bottom:1px solid #f1f5f9;font-size:13px;{vc}">{val}</td></tr>'
            )

        summary_rows = (
            tr("Client ID",     report["client_id"]) +
            tr("Client Name",   name) +
            tr("Period",        f"{start_lbl} – {end_lbl}") +
            tr("Total SMS",     f"{total:,}") +
            tr("Daily Average", f"{avg:,}") +
            tr("Peak Day",      f"{peak['date']} ({peak['count']:,} SMS)"     if peak   else "—", "#30318C") +
            tr("Lowest Day",    f"{lowest['date']} ({lowest['count']:,} SMS)" if lowest else "—")
        )

        html = (
            f'<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f0f2f8;font-family:Arial,sans-serif;">'
            f'<div style="max-width:600px;margin:32px auto;background:#fff;border-radius:10px;'
            f'overflow:hidden;box-shadow:0 4px 20px rgba(35,41,96,0.12);">'
            f'<div style="background:#ffffff;padding:18px 28px;border-bottom:3px solid #00B5A5;">'
            f'<img src="cid:vams_logo" alt="VAMS Global" style="height:46px;display:block;" /></div>'
            f'<div style="background:#30318C;padding:20px 28px;">'
            f'<div style="font-size:18px;color:#fff;font-weight:700;">Client Detailed SMS Report</div>'
            f'<div style="color:rgba(255,255,255,0.8);font-size:13px;margin-top:4px;">{name} &middot; {start_lbl} – {end_lbl}</div></div>'
            f'<div style="padding:20px 28px;">'
            f'<table style="width:100%;border-collapse:collapse;">{summary_rows}</table>'
            f'<div style="margin-top:16px;color:#374151;font-size:13px;">'
            f'The attached Excel file lists every individual SMS sent in this period '
            f'({sms_count:,} messages) — sender, mobile number, message text, and sent time.</div>'
            f'<div style="margin-top:16px;color:#94a3b8;font-size:11px;">Generated at {timestamp}</div>'
            f'</div>'
            f'<div style="background:#f8fafc;padding:12px 28px;border-top:1px solid #e8ecf4;">'
            f'<div style="font-size:11px;color:#64748b;"><strong style="color:#232960;">VAMS Global</strong> &middot; VAMS SMS Monitor</div>'
            f'<div style="font-size:10px;color:#94a3b8;margin-top:3px;">This is an automated report. Do not reply to this email.</div>'
            f'</div></div></body></html>'
        )

        plain = (
            f"VAMS Global  |  VAMS SMS Monitor\n"
            f"Client Detailed SMS Report — {name}  |  {start_lbl} – {end_lbl}\n"
            f"{'─'*48}\n"
            f"Client ID     : {report['client_id']}\n"
            f"Client Name   : {name}\n"
            f"Period        : {start_lbl} – {end_lbl}\n"
            f"Total SMS     : {total:,}\n"
            f"Daily Average : {avg:,}\n"
            f"See attached Excel for the full per-message detail ({sms_count:,} messages).\n"
            f"{'─'*48}\n"
            f"Generated at {timestamp}\n"
            f"This is an automated report. Do not reply."
        )

        subject = f"VAMS SMS Monitor: Detailed Usage Report — {name} ({start_lbl} to {end_lbl})"
        msg = MIMEMultipart("mixed")
        msg["Subject"] = Header(subject, "utf-8")
        msg["From"]    = self.from_addr
        if to_emails:
            msg["To"] = ", ".join(to_emails)
            cc_addrs  = [a for a in self.to_addrs if a not in to_emails]
            if cc_addrs:
                msg["Cc"] = ", ".join(cc_addrs)
        else:
            msg["To"] = ", ".join(self.to_addrs)
            cc_addrs  = []
        self._stamp_headers(msg)

        related = MIMEMultipart("related")
        alt = MIMEMultipart("alternative")
        alt.attach(MIMEText(plain, "plain", "utf-8"))
        alt.attach(MIMEText(html,  "html",  "utf-8"))
        related.attach(alt)
        logo = self._make_logo_part()
        if logo:
            related.attach(logo)
        msg.attach(related)

        part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        part.set_payload(excel_buf.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
        msg.attach(part)

        try:
            self._send(msg)
            recipients = to_emails if to_emails else list(self.to_addrs)
            print(f"[EMAIL] Detailed usage report for {name} sent to {', '.join(recipients)} ({sms_count:,} SMS)")
            return {"ok": True, "emails": recipients, "sms_count": sms_count}
        except Exception as e:
            print(f"[EMAIL ERROR] send_detailed_usage_report_email: {e}")
            return {"ok": False, "reason": "smtp_error"}

    def send_daily_report(self):
        """
        Sends one consolidated end-of-day report — a single table listing every
        flagged client with all their details and alert tags in one row.
        Scheduled daily at 18:00 from main.py.
        """
        if not self._is_configured():
            print("[EMAIL] SMTP not configured — skipping daily report.")
            return False

        try:
            conn   = get_connection()
            cursor = conn.cursor()
            cursor.execute("""
                SELECT
                    m.ClientName,
                    c.UsagePercent,
                    c.AllocatedSMS,
                    c.UsedSMS,
                    c.RemainingSMS,
                    ISNULL(tod.SmsCount, 0) AS TodayUsage,
                    ISNULL(yes.SmsCount, 0) AS YestUsage,
                    ISNULL(avg7.AvgSms,  0) AS SevenDayAvg
                FROM ClientSmsCredit_Vtb c WITH (NOLOCK)
                INNER JOIN ClientMst_Vtb m WITH (NOLOCK) ON c.ClientId = m.ClientId
                LEFT JOIN ClientSmsUsageDaily_Vtb tod WITH (NOLOCK)
                    ON tod.ClientId  = c.ClientId
                   AND tod.UsageDate = CAST(GETDATE() AS DATE)
                LEFT JOIN ClientSmsUsageDaily_Vtb yes WITH (NOLOCK)
                    ON yes.ClientId  = c.ClientId
                   AND yes.UsageDate = CAST(DATEADD(DAY, -1, GETDATE()) AS DATE)
                LEFT JOIN (
                    SELECT ClientId, AVG(CAST(SmsCount AS FLOAT)) AS AvgSms
                    FROM ClientSmsUsageDaily_Vtb WITH (NOLOCK)
                    WHERE UsageDate >= CAST(DATEADD(DAY, -7, GETDATE()) AS DATE)
                      AND UsageDate <  CAST(GETDATE() AS DATE)
                    GROUP BY ClientId
                ) avg7 ON avg7.ClientId = c.ClientId
                ORDER BY c.UsagePercent DESC, m.ClientName
            """)
            rows = cursor.fetchall()
            cursor.close()
            conn.close()
        except Exception as e:
            print(f"[EMAIL ERROR] send_daily_report DB fetch failed: {e}")
            return False

        t = get_thresholds()

        # ── Build flagged client list ──────────────────────────────────────
        # Row bg priority: critical > spike > drop > warning
        ROW_BG = {
            "critical": "#fff5f5",
            "spike":    "#faf8ff",
            "drop":     "#f0fdfb",
            "warning":  "#fffcf0",
        }
        PILL = {
            "critical": 'background:#fef2f2;color:#dc2626;border:1px solid #fecaca',
            "warning":  'background:#fffbeb;color:#d97706;border:1px solid #fde68a',
            "spike":    'background:#f5f3ff;color:#7c3aed;border:1px solid #ddd6fe',
            "drop":     'background:#e6faf9;color:#00B5A5;border:1px solid #99e0da',
        }
        PILL_LABEL = {
            "critical": f"{t['critical']}% Threshold",
            "warning":  f"{t['warning']}% Threshold",
            "spike":    "Spike Up",
            "drop":     "Spike Down",
        }

        flagged = []
        for r in rows:
            name      = r[0].strip()
            pct       = round(float(r[1]), 1)
            allocated = int(r[2])
            used      = int(r[3])
            remaining = int(r[4])
            today     = int(r[5])
            yest      = int(r[6])
            avg       = round(float(r[7]), 1)

            spike = (
                (avg > 0 and today > avg * 2.0 and today > yest * 1.5)
                or (avg <= 10 and yest <= 10 and today >= 50)
            )
            drop = yest > 50 and avg > 20 and today < avg * 0.5

            alert_types = []
            if pct >= t["critical"]:   alert_types.append("critical")
            elif pct >= t["warning"]:  alert_types.append("warning")
            if spike:       alert_types.append("spike")
            if drop:        alert_types.append("drop")

            if not alert_types:
                continue

            # Pick row background from highest-priority alert
            bg = ROW_BG.get(alert_types[0], "#fff")

            flagged.append(dict(
                name=name, pct=pct, allocated=allocated, used=used,
                remaining=remaining, today=today, yest=yest, avg=avg,
                ratio=round(today / max(avg, 1), 1),
                alerts=alert_types, bg=bg
            ))

        n_clients  = len(flagged)
        report_date = datetime.now().strftime("%d %b %Y")
        gen_time    = datetime.now().strftime("%I:%M %p")

        subject_line = (
            f"VAMS SMS Monitor: Daily Report — {report_date} — All Clear"
            if n_clients == 0 else
            f"VAMS SMS Monitor: Daily Report — {report_date} — {n_clients} client{'s' if n_clients > 1 else ''} flagged"
        )

        # ── Summary counts ─────────────────────────────────────────────────
        n_crit  = sum(1 for c in flagged if "critical" in c["alerts"])
        n_warn  = sum(1 for c in flagged if "warning"  in c["alerts"])
        n_spike = sum(1 for c in flagged if "spike"    in c["alerts"])
        n_drop  = sum(1 for c in flagged if "drop"     in c["alerts"])

        def pill(style, label):
            return (f'<span style="display:inline-block;{style};padding:3px 10px;'
                    f'border-radius:100px;font-size:11px;font-weight:700;'
                    f'margin-right:6px;">{label}</span>')

        # ── Column header helper ───────────────────────────────────────────
        def th(text, align="left"):
            return (
                f'<th style="text-align:{align};padding:10px 12px;background:#30318C;'
                f'color:#fff;font-size:11px;font-weight:600;letter-spacing:0.5px;'
                f'text-transform:uppercase;white-space:nowrap;">{text}</th>'
            )

        # ── Cell helper ────────────────────────────────────────────────────
        def td(text, bold_color=None, align="left", bg=""):
            style = f"padding:9px 12px;font-size:13px;border-bottom:1px solid #f1f5f9;text-align:{align};"
            if bg:       style += f"background:{bg};"
            if bold_color: style += f"color:{bold_color};font-weight:700;"
            return f'<td style="{style}">{text}</td>'

        # ── Build table rows ───────────────────────────────────────────────
        if n_clients == 0:
            table_html = (
                f'<tr><td colspan="9" style="text-align:center;padding:40px;color:#64748b;font-size:14px;">'
                f'&#10003; All clients are healthy — no alerts today.'
                f'</td></tr>'
            )
        else:
            table_html = ""
            for i, c in enumerate(flagged, 1):
                bg = c["bg"]
                # alert pills for this row
                pills = "".join(pill(PILL[a], PILL_LABEL[a]) for a in c["alerts"])
                # usage % color
                pct_color = "#dc2626" if c["pct"] >= t["critical"] else "#d97706" if c["pct"] >= t["warning"] else "#16a34a"

                table_html += (
                    f'<tr>'
                    + td(str(i),              align="center",  bg=bg)
                    + td(f'<strong>{c["name"]}</strong>', bg=bg)
                    + td(f'{c["allocated"]:,}', align="right",  bg=bg)
                    + td(f'{c["used"]:,}',       align="right",  bg=bg)
                    + td(f'{c["remaining"]:,}',  align="right",  bg=bg, bold_color="#16a34a")
                    + td(f'{c["pct"]}%',          align="right",  bg=bg, bold_color=pct_color)
                    + td(f'{c["today"]:,}',       align="right",  bg=bg)
                    + td(f'{c["avg"]:,}',         align="right",  bg=bg)
                    + td(pills,                   bg=bg)
                    + f'</tr>'
                )

        # ── Summary bar ────────────────────────────────────────────────────
        summary_pills = ""
        if n_crit:  summary_pills += pill(PILL["critical"], f"{n_crit} Critical")
        if n_warn:  summary_pills += pill(PILL["warning"],  f"{n_warn} Warning")
        if n_spike: summary_pills += pill(PILL["spike"],    f"{n_spike} Spike")
        if n_drop:  summary_pills += pill(PILL["drop"],     f"{n_drop} Drop")
        if not summary_pills:
            summary_pills = '<span style="color:#16a34a;font-weight:600;">&#10003; No alerts today</span>'

        # ── Full HTML ──────────────────────────────────────────────────────
        html = (
            f'<!DOCTYPE html><html>'
            f'<body style="margin:0;padding:0;background:#f0f2f8;font-family:Arial,sans-serif;">'
            f'<div style="max-width:820px;margin:32px auto;background:#fff;border-radius:10px;'
            f'overflow:hidden;box-shadow:0 4px 20px rgba(35,41,96,0.12);">'

            # VAMS brand bar
            f'<div style="background:#232960;padding:13px 28px;display:flex;align-items:center;gap:10px;">'
            f'<div style="width:3px;height:28px;background:#00B5A5;border-radius:2px;flex-shrink:0;"></div>'
            f'<div>'
            f'<div style="font-size:10px;font-weight:700;color:#00B5A5;letter-spacing:2px;text-transform:uppercase;">VAMS Global</div>'
            f'<div style="font-size:13px;font-weight:700;color:#ffffff;letter-spacing:0.5px;">VAMS SMS Monitor</div>'
            f'</div>'
            f'</div>'

            # Report header (table layout — flex is unreliable in email clients)
            f'<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#30318C;">'
            f'<tr>'
            f'<td style="padding:20px 28px;vertical-align:middle;">'
            f'<div style="font-size:18px;color:#fff;font-weight:700;">Daily Alert Report</div>'
            f'<div style="color:rgba(255,255,255,0.78);font-size:12px;margin-top:3px;">{report_date} &middot; Generated at {gen_time}</div>'
            f'</td>'
            f'<td style="padding:20px 28px;vertical-align:middle;text-align:right;white-space:nowrap;width:90px;">'
            f'<span style="background:rgba(255,255,255,0.15);color:#fff;font-size:22px;font-weight:800;'
            f'padding:8px 18px;border-radius:8px;border:1px solid rgba(255,255,255,0.2);display:inline-block;">{n_clients}</span>'
            f'</td>'
            f'</tr>'
            f'</table>'

            # Summary bar
            f'<div style="padding:14px 28px;border-bottom:1px solid #e8ecf4;background:#f8fafc;">'
            f'{summary_pills}'
            f'</div>'

            # Table
            f'<div style="overflow-x:auto;">'
            f'<table style="width:100%;border-collapse:collapse;min-width:700px;">'
            f'<thead><tr>'
            f'{th("#", "center")}{th("Client")}{th("Allocated", "right")}'
            f'{th("Used", "right")}{th("Remaining", "right")}{th("Usage %", "right")}'
            f'{th("Today", "right")}{th("7-Day Avg", "right")}{th("Alerts")}'
            f'</tr></thead>'
            f'<tbody>{table_html}</tbody>'
            f'</table>'
            f'</div>'

            # VAMS footer
            f'<div style="background:#f8fafc;padding:12px 28px;border-top:1px solid #e8ecf4;">'
            f'<div style="font-size:11px;color:#64748b;">'
            f'<strong style="color:#232960;">VAMS Global</strong>'
            f'&nbsp;&middot;&nbsp;VAMS SMS Monitor &nbsp;&middot;&nbsp; Automated Daily Report'
            f'</div>'
            f'<div style="font-size:10px;color:#94a3b8;margin-top:3px;">This is an automated notification. Do not reply to this email.</div>'
            f'</div>'

            f'</div></body></html>'
        )

        # ── Build Excel attachment (2 sheets) ─────────────────────────────
        thin   = Side(style="thin", color="D1D5DB")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        ALERT_LABEL = {
            "critical": f"{t['critical']}% Threshold",
            "warning":  f"{t['warning']}% Threshold",
            "spike":    "Spike Up",
            "drop":     "Spike Down",
        }
        ROW_FILL = {
            "critical": PatternFill("solid", fgColor="FFF5F5"),
            "warning":  PatternFill("solid", fgColor="FFFCF0"),
            "spike":    PatternFill("solid", fgColor="FAF8FF"),
            "drop":     PatternFill("solid", fgColor="F0FDFB"),
        }
        SECTION_COLOR = {
            "critical": "DC2626",
            "warning":  "D97706",
            "spike":    "7C3AED",
            "drop":     "00B5A5",
        }

        def _write_section_sheet(ws, sections):
            headers    = ["#", "Client", "Allocated", "Used", "Remaining",
                          "Usage %", "Today", "7-Day Avg", "Alert Types"]
            col_widths = [5, 30, 14, 14, 14, 10, 12, 12, 36]

            for ci, w in enumerate(col_widths, 1):
                ws.column_dimensions[get_column_letter(ci)].width = w

            # ── Row 1: VAMS brand bar ──────────────────────────────────────
            ws.merge_cells("A1:I1")
            brand = ws.cell(row=1, column=1,
                            value=f"VAMS Global  |  VAMS SMS Monitor  |  Daily Alert Report  |  {report_date}")
            brand.font      = Font(name="Calibri", bold=True, color="FFFFFF", size=12)
            brand.fill      = PatternFill("solid", fgColor="232960")
            brand.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            ws.row_dimensions[1].height = 24

            # Teal left-border accent on brand row
            from openpyxl.styles import Side as _Side, Border as _Border
            teal_left = _Border(left=_Side(style="thick", color="00B5A5"))
            brand.border = teal_left

            # ── Row 2: Column headers ──────────────────────────────────────
            for ci, h in enumerate(headers, 1):
                cell = ws.cell(row=2, column=ci, value=h)
                cell.font      = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
                cell.fill      = PatternFill("solid", fgColor="30318C")
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border    = border
            ws.row_dimensions[2].height = 22
            ws.freeze_panes = "B3"

            data_row = 3
            seq = 0
            for section_key, clients in sections:
                if not clients:
                    continue
                n = len(clients)
                ws.merge_cells(f"A{data_row}:I{data_row}")
                sh = ws.cell(row=data_row, column=1)
                sh.value     = f"  {ALERT_LABEL[section_key]}  —  {n} client{'s' if n != 1 else ''}"
                sh.font      = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
                sh.fill      = PatternFill("solid", fgColor=SECTION_COLOR[section_key])
                sh.alignment = Alignment(horizontal="left", vertical="center")
                ws.row_dimensions[data_row].height = 22
                data_row += 1

                for client in clients:
                    seq      += 1
                    rf        = ROW_FILL[section_key]
                    alert_str = ", ".join(ALERT_LABEL[a] for a in client["alerts"])
                    pct_color = "DC2626" if client["pct"] >= t["critical"] else "D97706" if client["pct"] >= t["warning"] else "166534"

                    values = [seq, client["name"], client["allocated"], client["used"],
                              client["remaining"], client["pct"], client["today"], client["avg"], alert_str]
                    aligns = ["center", "left", "right", "right",
                              "right", "right", "right", "right", "left"]

                    for ci, (val, aln) in enumerate(zip(values, aligns), 1):
                        cell = ws.cell(row=data_row, column=ci, value=val)
                        cell.fill      = rf
                        cell.border    = border
                        cell.alignment = Alignment(horizontal=aln, vertical="center")
                        if ci == 6:
                            cell.font = Font(name="Calibri", bold=True, color=pct_color, size=10)
                        else:
                            cell.font = Font(name="Calibri", size=10)
                    ws.row_dimensions[data_row].height = 18
                    data_row += 1

        # Each sheet shows all clients matching that category
        # (a client at 95% with a spike appears in both sheets)
        spike_clients = [c for c in flagged if "spike" in c["alerts"]]
        drop_clients  = [c for c in flagged if "drop"  in c["alerts"]]
        crit_clients  = [c for c in flagged if "critical" in c["alerts"]]
        warn_clients  = [c for c in flagged if "warning"  in c["alerts"]]

        wb = openpyxl.Workbook()

        ws_spikes       = wb.active
        ws_spikes.title = "Spikes"
        _write_section_sheet(ws_spikes, [("spike", spike_clients), ("drop", drop_clients)])

        ws_thresh       = wb.create_sheet("Thresholds")
        _write_section_sheet(ws_thresh, [("critical", crit_clients), ("warning", warn_clients)])

        excel_buf = io.BytesIO()
        wb.save(excel_buf)
        excel_buf.seek(0)

        filename = f"sms_alert_report_{datetime.now().strftime('%Y%m%d')}.xlsx"

        # ── Assemble email ─────────────────────────────────────────────────
        msg = MIMEMultipart("mixed")
        msg["Subject"] = Header(subject_line, "utf-8")
        msg["From"]    = self.from_addr
        msg["To"]      = ", ".join(self.to_addrs)
        self._stamp_headers(msg)

        plain_text = (
            f"VAMS Global  |  VAMS SMS Monitor\n"
            f"Daily Alert Report — {report_date}\n"
            f"{'─' * 44}\n"
            f"Clients flagged : {n_clients}\n"
            f"Critical        : {n_crit}\n"
            f"Warning         : {n_warn}\n"
            f"Spike           : {n_spike}\n"
            f"Drop            : {n_drop}\n"
            f"{'─' * 44}\n"
            f"See the attached Excel file for full details.\n"
            f"Generated at {gen_time}\n"
            f"This is an automated notification. Do not reply."
        )
        alt = MIMEMultipart("alternative")
        alt.attach(MIMEText(plain_text, "plain", "utf-8"))
        alt.attach(MIMEText(html,       "html",  "utf-8"))
        msg.attach(alt)

        part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        part.set_payload(excel_buf.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
        msg.attach(part)

        try:
            srv = self._smtp_connect()
            if srv.sock:
                srv.sock.settimeout(60)
            try:
                srv.send_message(msg)
            finally:
                try:
                    srv.quit()
                except Exception:
                    pass
            print(f"[EMAIL] Daily report sent to {', '.join(self.to_addrs)} ({n_clients} client(s) flagged) + Excel attachment")
            return True
        except Exception as e:
            print(f"[EMAIL ERROR] Daily report failed: {e}")
            return False


notifier = EmailNotifier()


def fetch_client_analytics(client_id):
    """Credit summary + last 7 days of daily usage for a single client. Cached 60s."""
    cached = _dcache_analytics.get(client_id)
    if cached and (datetime.now() - cached["ts"]).total_seconds() < _CACHE_TTL:
        return cached["data"]
    try:
        conn   = get_connection()
        cursor = conn.cursor()

        cursor.execute("""
            SELECT c.AllocatedSMS, c.UsedSMS, c.RemainingSMS, c.UsagePercent,
                   m.ClientName
            FROM ClientSmsCredit_Vtb c
            INNER JOIN ClientMst_Vtb m ON c.ClientId = m.ClientId
            WHERE c.ClientId = ?
        """, client_id)
        row = cursor.fetchone()
        if not row:
            cursor.close()
            conn.close()
            return None

        cursor.execute("""
            SELECT UsageDate, SmsCount
            FROM ClientSmsUsageDaily_Vtb
            WHERE ClientId = ?
              AND UsageDate >= CAST(DATEADD(DAY, -7, GETDATE()) AS DATE)
            ORDER BY UsageDate
        """, client_id)
        daily_rows = cursor.fetchall()

        cursor.close()
        conn.close()

        daily = [{"date": r[0].strftime("%d %b"), "count": int(r[1])} for r in daily_rows]

        today_d = date.today()
        yest_d  = today_d - timedelta(days=1)
        today_count   = next((int(r[1]) for r in daily_rows if r[0] == today_d), 0)
        yest_count    = next((int(r[1]) for r in daily_rows if r[0] == yest_d),  0)
        past_days     = [int(r[1]) for r in daily_rows if r[0] != today_d]
        seven_day_avg = round(sum(past_days) / len(past_days), 1) if past_days else 0.0

        spike_detected = (
            (seven_day_avg > 0 and today_count > seven_day_avg * 2.0 and today_count > yest_count * 1.5)
            or
            (seven_day_avg <= 10 and yest_count <= 10 and today_count >= 50)
        )
        drop_detected = (
            yest_count > 50 and
            seven_day_avg > 20 and
            today_count < seven_day_avg * 0.5
        )

        result = {
            "client_id":      client_id,
            "client_name":    row[4].strip(),
            "allocated":      int(row[0]),
            "used":           int(row[1]),
            "remaining":      int(row[2]),
            "usage_pct":      round(float(row[3]), 1),
            "today_count":    today_count,
            "yest_count":     yest_count,
            "seven_day_avg":  seven_day_avg,
            "spike_detected": spike_detected,
            "drop_detected":  drop_detected,
            "daily":          daily
        }
        _dcache_analytics[client_id] = {"data": result, "ts": datetime.now()}
        return result
    except Exception as e:
        print(f"[DB ERROR] fetch_client_analytics: {e}")
        return None


def cleanup_orphaned_rows() -> dict:
    """
    Deletes rows in ClientSmsCredit_Vtb and ClientSmsUsageDaily_Vtb whose
    ClientId no longer exists in ClientMst_Vtb. Called after every hourly sync
    so stale data never accumulates. Uses NOT EXISTS (not NOT IN) to handle
    NULLs correctly and perform well on large tables.
    """
    try:
        conn   = get_connection()
        cursor = conn.cursor()

        cursor.execute("""
            DELETE FROM ClientSmsCredit_Vtb
            WHERE NOT EXISTS (
                SELECT 1 FROM ClientMst_Vtb m
                WHERE m.ClientId = ClientSmsCredit_Vtb.ClientId
            )
        """)
        credits_removed = cursor.rowcount

        cursor.execute("""
            DELETE FROM ClientSmsUsageDaily_Vtb
            WHERE NOT EXISTS (
                SELECT 1 FROM ClientMst_Vtb m
                WHERE m.ClientId = ClientSmsUsageDaily_Vtb.ClientId
            )
        """)
        daily_removed = cursor.rowcount

        conn.commit()
        cursor.close()
        conn.close()

        if credits_removed or daily_removed:
            print(f"[CLEANUP] Removed {credits_removed} orphaned credit row(s), "
                  f"{daily_removed} orphaned daily row(s).")
        return {"credits_removed": credits_removed, "daily_removed": daily_removed}

    except Exception as e:
        print(f"[DB ERROR] cleanup_orphaned_rows: {e}")
        return {"credits_removed": 0, "daily_removed": 0}


def sync_sms_usage():
    """
    Syncs real SMS usage from Smslog_Vtb into ClientSmsUsageDaily_Vtb and ClientSmsCredit_Vtb.

    Processes the last 30 days so historical data is backfilled on first run,
    and today's data is always kept current on hourly runs.

    Flow:
      1. Aggregate SMS per client per day from Smslog_Vtb (last 30 days)
      2. Upsert each day's count into ClientSmsUsageDaily_Vtb
      3. For each affected client, recalculate total UsedSMS from ClientSmsUsageDaily_Vtb
      4. Update ClientSmsCredit_Vtb (UsedSMS, RemainingSMS, UsagePercent)
    """
    try:
        conn   = get_connection()
        cursor = conn.cursor()

        # ── Step 1+2: Aggregate from Smslog_Vtb and upsert ──────────────────
        cursor.execute("""
            MERGE ClientSmsUsageDaily_Vtb AS target
            USING (
                SELECT
                    s.senderCompId                                                      AS ClientId,
                    CAST(s.SentTime AS DATE)                                            AS UsageDate,
                    SUM(CEILING(LEN(LTRIM(RTRIM(ISNULL(s.Message, '')))) / 160.0))     AS SmsCount
                FROM Smslog_Vtb s
                INNER JOIN ClientSmsCredit_Vtb c ON s.senderCompId = c.ClientId
                WHERE CAST(s.SentTime AS DATE) >= DATEADD(DAY, -30, CAST(GETDATE() AS DATE))
                  AND s.senderCompId IS NOT NULL
                GROUP BY s.senderCompId, CAST(s.SentTime AS DATE)
            ) AS source
            ON  target.ClientId  = source.ClientId
            AND target.UsageDate = source.UsageDate
            WHEN MATCHED THEN
                UPDATE SET target.SmsCount = source.SmsCount
            WHEN NOT MATCHED THEN
                INSERT (ClientId, UsageDate, SmsCount, CreatedDate)
                VALUES (source.ClientId, source.UsageDate, source.SmsCount, GETDATE());
        """)
        synced_rows = cursor.rowcount

        # ── Step 3+4: Recalculate totals scoped to each client's validity window ─
        # For clients with ValidityStartDate: sum only within [StartDate, EndDate].
        # For all others: default 30-day window.
        cursor.execute("""
            ;WITH usage_scoped AS (
                SELECT
                    c.ClientId,
                    ISNULL(SUM(d.SmsCount), 0) AS ScopedUsed
                FROM ClientSmsCredit_Vtb c
                LEFT JOIN ClientSmsUsageDaily_Vtb d
                    ON d.ClientId  = c.ClientId
                   AND d.UsageDate >= ISNULL(c.ValidityStartDate,
                                             CAST(DATEADD(DAY, -30, GETDATE()) AS DATE))
                   AND d.UsageDate <= CASE
                           WHEN c.ValidityEndDate IS NOT NULL
                            AND c.ValidityEndDate < CAST(GETDATE() AS DATE)
                           THEN c.ValidityEndDate
                           ELSE CAST(GETDATE() AS DATE)
                       END
                GROUP BY c.ClientId
            )
            UPDATE c
            SET
                c.UsedSMS      = u.ScopedUsed,
                c.RemainingSMS = CASE
                    WHEN c.AllocatedSMS > u.ScopedUsed
                    THEN c.AllocatedSMS - u.ScopedUsed
                    ELSE 0
                END,
                c.ModifiedDate = GETDATE()
            FROM ClientSmsCredit_Vtb c
            INNER JOIN usage_scoped u ON c.ClientId = u.ClientId
        """)
        affected_clients = cursor.rowcount

        conn.commit()
        cursor.close()
        conn.close()

        global _last_synced_at
        _last_synced_at = datetime.now()
        invalidate_dashboard_cache()   # force fresh data on next load

        print(f"[SYNC] {synced_rows} day-rows upserted, {affected_clients} client(s) recalculated.")
        cleanup_orphaned_rows()
        return {"synced_rows": synced_rows, "affected_clients": affected_clients}

    except Exception as e:
        import traceback
        print("[ERROR] sync_sms_usage failed:")
        traceback.print_exc()
        return None


def seed_test_data():
    """
    Creates test SMS usage data for all real clients in ClientMst_Vtb.
    Inserts directly into ClientSmsCredit_Vtb and ClientSmsUsageDaily_Vtb.

    14 scenarios cycle across clients — majority are spike/drop-ONLY so the
    anomaly section is always well-populated:

    Spike-ONLY (6):
      today = 8× daily_base  →  today > avg×2 AND today > yest×1.5  → spike flag set
      usage_pct < 20%        →  well below 50% warning threshold     → no threshold alert

    Drop-ONLY (5):
      yesterday = 4× daily_base, today = daily_base÷8
                             →  today < avg×0.5 AND yest>50 AND avg>20 → drop flag set
      usage_pct < 35%        →  well below 50% warning threshold     → no threshold alert

    Normal (3 — for a complete dashboard):
      - Healthy  (25% usage)
      - Warning  (55% usage)
      - Critical (92% usage)
    """
    scenarios = [
        # ── Spike-ONLY (usage ≈ 8-18%, no threshold triggered) ──────────────
        # 7-day avg ≈ daily_base (today excluded from avg)
        # today = 8×avg → 8>2 ✓   today = 8×yest → 8>1.5 ✓
        {"allocated": 1000, "target_pct": 0.08, "pattern": "spike"},
        {"allocated": 800,  "target_pct": 0.10, "pattern": "spike"},
        {"allocated": 700,  "target_pct": 0.12, "pattern": "spike"},
        {"allocated": 600,  "target_pct": 0.14, "pattern": "spike"},
        {"allocated": 500,  "target_pct": 0.16, "pattern": "spike"},
        {"allocated": 450,  "target_pct": 0.18, "pattern": "spike"},
        # ── Drop-ONLY  (usage ≈ 15-33%, no threshold triggered) ─────────────
        # avg ≈ 1.5×daily_base (yesterday = 4× inflates avg)
        # today = base÷8 < avg×0.5 ✓   yest = 4×base > 50 ✓   avg > 20 ✓
        {"allocated": 900,  "target_pct": 0.15, "pattern": "drop"},
        {"allocated": 750,  "target_pct": 0.18, "pattern": "drop"},
        {"allocated": 650,  "target_pct": 0.20, "pattern": "drop"},
        {"allocated": 550,  "target_pct": 0.22, "pattern": "drop"},
        {"allocated": 400,  "target_pct": 0.25, "pattern": "drop"},
        # ── Normal (threshold coverage) ──────────────────────────────────────
        {"allocated": 500,  "target_pct": 0.25, "pattern": "normal"},   # healthy
        {"allocated": 300,  "target_pct": 0.55, "pattern": "normal"},   # warning
        {"allocated": 150,  "target_pct": 0.92, "pattern": "normal"},   # critical
    ]

    today = date.today()

    try:
        conn   = get_connection()
        cursor = conn.cursor()

        # Seed all clients — patterns are maintained hourly by refresh_hourly_test_usage
        # (called automatically at the end of every sync_sms_usage run).
        cursor.execute("""
            SELECT ClientId, ClientName
            FROM ClientMst_Vtb
            WHERE ClientName IS NOT NULL
              AND LTRIM(RTRIM(ClientName)) != ''
            ORDER BY ClientName
        """)
        clients = cursor.fetchall()

        if not clients:
            print("[SEED] No clients found in ClientMst_Vtb.")
            return {"seeded_clients": 0}

        # Clients with ManualAlloc=1 have admin-assigned credits — skip them entirely
        cursor.execute("SELECT ClientId FROM ClientSmsCredit_Vtb WITH (NOLOCK) WHERE ManualAlloc = 1")
        manual_ids = {row[0] for row in cursor.fetchall()}

        # ── Build all data in Python first (no DB round-trips in loop) ──────
        credit_rows = []   # (allocated, used, remaining, client_id)
        daily_rows  = []   # (client_id, day_date, sms_count)

        for i, client_row in enumerate(clients):
            client_id   = client_row[0]
            if client_id in manual_ids:
                continue
            scenario    = scenarios[i % len(scenarios)]
            allocated   = scenario["allocated"]
            target_used = int(allocated * scenario["target_pct"])
            pattern     = scenario["pattern"]

            days       = 7
            daily_base = max(1, target_used // days)

            day_counts = []
            for d in range(days - 1, -1, -1):   # index 6 = oldest, 0 = today
                day_date = today - timedelta(days=d)
                if d == 0 and pattern == "spike":
                    count = daily_base * 8
                elif d == 1 and pattern == "spike":
                    count = daily_base
                elif d == 0 and pattern == "drop":
                    count = max(1, daily_base // 8)
                elif d == 1 and pattern == "drop":
                    count = daily_base * 4
                else:
                    count = int(daily_base * random.uniform(0.85, 1.15))
                day_counts.append((day_date, max(1, count)))

            actual_total = sum(c for _, c in day_counts)
            remaining    = max(0, allocated - actual_total)

            credit_rows.append((allocated, actual_total, remaining, client_id))
            for day_date, sms_count in day_counts:
                daily_rows.append((client_id, day_date, sms_count))

        seeded = len(credit_rows)
        skipped = len(clients) - seeded
        print(f"[SEED] Built data for {seeded} clients ({len(daily_rows)} daily rows), skipped {skipped} manually-allocated. Writing to DB...")

        cursor.fast_executemany = True   # use SQL Server batch arrays — much faster

        # ── Bulk-upsert credits (single MERGE per row via executemany) ───────
        cursor.executemany("""
            MERGE ClientSmsCredit_Vtb AS t
            USING (SELECT ? AS alloc, ? AS used, ? AS rem, ? AS cid)
                  AS s(alloc, used, rem, cid)
               ON t.ClientId = s.cid
            WHEN MATCHED THEN
                UPDATE SET AllocatedSMS = s.alloc, UsedSMS        = s.used,
                           RemainingSMS = s.rem,   Alert50Sent    = 0,
                           Alert90Sent  = 0,       AlertSpikeSent = 0,
                           AlertDropSent = 0,      ManualAlloc    = 0,
                           ModifiedDate = GETDATE()
            WHEN NOT MATCHED THEN
                INSERT (ClientId, AllocatedSMS, UsedSMS, RemainingSMS,
                        Alert50Sent, Alert90Sent, ManualAlloc,
                        CreatedDate, ModifiedDate)
                VALUES (s.cid, s.alloc, s.used, s.rem,
                        0, 0, 0, GETDATE(), GETDATE());
        """, credit_rows)
        print(f"[SEED] Credits upserted.")

        # ── Bulk-replace daily rows: delete per-client then batch insert ────
        # Delete only the seeded clients' last-7-days rows (avoids full-table lock)
        client_ids = [r[3] for r in credit_rows]
        chunk_size = 200
        for i in range(0, len(client_ids), chunk_size):
            chunk = client_ids[i:i + chunk_size]
            placeholders = ",".join("?" * len(chunk))
            cursor.execute(f"""
                DELETE FROM ClientSmsUsageDaily_Vtb
                WHERE ClientId IN ({placeholders})
                  AND UsageDate >= CAST(DATEADD(DAY, -7, GETDATE()) AS DATE)
            """, chunk)

        # Insert daily rows in batches to avoid connection timeout
        for i in range(0, len(daily_rows), chunk_size):
            cursor.executemany("""
                INSERT INTO ClientSmsUsageDaily_Vtb (ClientId, UsageDate, SmsCount, CreatedDate)
                VALUES (?, ?, ?, GETDATE())
            """, daily_rows[i:i + chunk_size])

        print(f"[SEED] Daily rows inserted.")

        # ── Final step: set AllocatedSMS per slot for a realistic threshold mix ──
        #
        #  slot 0-5  (spike)  → AllocatedSMS = UsedSMS × 7  → ~14% → healthy
        #  slot 6-10 (drop)   → AllocatedSMS = UsedSMS × 5  → ~20% → healthy
        #  slot 11   (normal) → AllocatedSMS = UsedSMS × 3  → ~33% → healthy
        #  slot 12   (normal) → AllocatedSMS = UsedSMS / 65%         → warning
        #  slot 13   (normal) → AllocatedSMS = UsedSMS / 95%         → critical
        cursor.execute("""
            ;WITH ranked AS (
                SELECT ClientId,
                       (ROW_NUMBER() OVER (ORDER BY ClientId) - 1) % 14 AS slot
                FROM ClientSmsCredit_Vtb WITH (NOLOCK)
            ),
            targets AS (
                SELECT c.ClientId,
                       CASE
                           WHEN r.slot < 6
                               THEN CASE WHEN c.UsedSMS > 0 THEN c.UsedSMS * 7  ELSE 5000 END
                           WHEN r.slot < 11
                               THEN CASE WHEN c.UsedSMS > 0 THEN c.UsedSMS * 5  ELSE 3000 END
                           WHEN r.slot = 11
                               THEN CASE WHEN c.UsedSMS > 0 THEN c.UsedSMS * 3  ELSE 2000 END
                           WHEN r.slot = 12
                               THEN CASE WHEN c.UsedSMS > 0
                                         THEN CEILING(CAST(c.UsedSMS AS FLOAT) / 0.65)
                                         ELSE 1000 END
                           ELSE  CASE WHEN c.UsedSMS > 0
                                      THEN CEILING(CAST(c.UsedSMS AS FLOAT) / 0.95)
                                      ELSE 500  END
                       END AS new_alloc
                FROM ClientSmsCredit_Vtb c WITH (NOLOCK)
                INNER JOIN ranked r ON r.ClientId = c.ClientId
            )
            UPDATE c
            SET c.AllocatedSMS = t.new_alloc,
                c.RemainingSMS = CASE WHEN t.new_alloc > c.UsedSMS
                                      THEN t.new_alloc - c.UsedSMS ELSE 0 END,
                c.ModifiedDate = GETDATE()
            FROM ClientSmsCredit_Vtb c
            INNER JOIN targets t ON t.ClientId = c.ClientId
            WHERE c.ManualAlloc = 0
        """)

        conn.commit()
        cursor.close()
        conn.close()

        invalidate_dashboard_cache()   # show seeded data on next dashboard load
        print(f"[SEED] Done. {seeded} clients seeded.")
        return {"seeded_clients": seeded}

    except Exception as e:
        import traceback
        print("[ERROR] seed_test_data failed:")
        traceback.print_exc()
        return None


def ensure_today_yesterday_data():
    """
    Called once on startup. Bulk-inserts today and yesterday rows for every client
    that is missing them — pure set-based SQL, no Python loops, runs in milliseconds.

    Uses the client's own historical average as the seed value so numbers look
    realistic from the first page load. Falls back to 500/300 for brand-new clients.
    Also recalculates ClientSmsCredit_Vtb totals for any client that got new rows.
    """
    try:
        conn   = get_connection()
        cursor = conn.cursor()

        # ── Insert yesterday for clients missing it ────────────────────
        cursor.execute("""
            INSERT INTO ClientSmsUsageDaily_Vtb (ClientId, UsageDate, SmsCount, CreatedDate)
            SELECT
                c.ClientId,
                CAST(DATEADD(DAY, -1, GETDATE()) AS DATE),
                ISNULL((
                    SELECT AVG(SmsCount)
                    FROM ClientSmsUsageDaily_Vtb h WITH (NOLOCK)
                    WHERE h.ClientId = c.ClientId
                ), 0),
                GETDATE()
            FROM ClientSmsCredit_Vtb c WITH (NOLOCK)
            WHERE NOT EXISTS (
                SELECT 1 FROM ClientSmsUsageDaily_Vtb d WITH (NOLOCK)
                WHERE d.ClientId  = c.ClientId
                  AND d.UsageDate = CAST(DATEADD(DAY, -1, GETDATE()) AS DATE)
            )
        """)
        yesterday_inserted = cursor.rowcount

        # ── Insert today for clients missing it ────────────────────────
        cursor.execute("""
            INSERT INTO ClientSmsUsageDaily_Vtb (ClientId, UsageDate, SmsCount, CreatedDate)
            SELECT
                c.ClientId,
                CAST(GETDATE() AS DATE),
                ISNULL((
                    SELECT AVG(SmsCount)
                    FROM ClientSmsUsageDaily_Vtb h WITH (NOLOCK)
                    WHERE h.ClientId  = c.ClientId
                      AND h.UsageDate < CAST(GETDATE() AS DATE)
                ), 0),
                GETDATE()
            FROM ClientSmsCredit_Vtb c WITH (NOLOCK)
            WHERE NOT EXISTS (
                SELECT 1 FROM ClientSmsUsageDaily_Vtb d WITH (NOLOCK)
                WHERE d.ClientId  = c.ClientId
                  AND d.UsageDate = CAST(GETDATE() AS DATE)
            )
        """)
        today_inserted = cursor.rowcount

        # ── Recalculate ClientSmsCredit_Vtb totals for all clients ─────
        # Scoped to last 30 days (matches sync_sms_usage window) so all-time
        # accumulation can never inflate UsagePercent beyond column precision.
        cursor.execute("""
            UPDATE c
            SET
                c.UsedSMS      = agg.TotalUsed,
                c.RemainingSMS = CASE
                                    WHEN c.AllocatedSMS > agg.TotalUsed
                                    THEN c.AllocatedSMS - agg.TotalUsed
                                    ELSE 0
                                 END,
                c.ModifiedDate = GETDATE()
            FROM ClientSmsCredit_Vtb c
            INNER JOIN (
                SELECT ClientId, SUM(SmsCount) AS TotalUsed
                FROM ClientSmsUsageDaily_Vtb
                WHERE UsageDate >= CAST(DATEADD(DAY, -30, GETDATE()) AS DATE)
                GROUP BY ClientId
            ) agg ON agg.ClientId = c.ClientId
        """)

        conn.commit()
        cursor.close()
        conn.close()
        print(f"[STARTUP] ensure_today_yesterday_data: +{yesterday_inserted} yesterday rows, +{today_inserted} today rows.")

    except Exception as e:
        print(f"[ERROR] ensure_today_yesterday_data: {e}")


def refresh_hourly_test_usage():
    """
    Called automatically at the end of every sync_sms_usage() run and also via
    the manual /api/refresh-test-data endpoint.

    Re-applies deterministic spike/drop/normal patterns to ALL clients so the
    dashboard always has visible anomaly alerts regardless of real Smslog data.

    Pattern is assigned by client position (sorted by ClientId, index % 14):
      slots  0-5  → spike  : days -7 to -1 = 50, today = 500
                              (today(500) > avg7(50)×2=100 ✓  today > yest(50)×1.5=75 ✓)
      slots  6-10 → drop   : days -7 to -1 = 50, yesterday = 250, today = 3
                              (today(3) < avg7(~79)×0.5 ✓  yest(250)>50 ✓  avg>20 ✓)
      slots 11-13 → normal : left untouched (shows real sync data)

    Historical days are normalised to a fixed baseline (DAILY_BASE=50) before
    writing today/yesterday so that spike/drop detection — which derives avg7
    from those rows — always sees bounded values regardless of real production
    volumes.  Without this, real Smslog data (thousands/day) fed through the
    old avg7×10 / avg7×5 multipliers produced unrealistically large numbers.
    """
    # ranked: assigns each client a slot (0-13) based on ClientId sort order
    _CTE = """
        ;WITH ranked AS (
            SELECT ClientId,
                   (ROW_NUMBER() OVER (ORDER BY ClientId) - 1) % 14 AS slot
            FROM ClientSmsCredit_Vtb WITH (NOLOCK)
        )
    """

    try:
        conn   = get_connection()
        cursor = conn.cursor()

        # ── 0. Normalise history: set days -7 to -1 = 50 for pattern clients ─
        # Resets any large real-data values written by sync_sms_usage so the
        # avg7 seen by spike/drop detection is always bounded.
        cursor.execute(_CTE + """
            MERGE ClientSmsUsageDaily_Vtb AS tgt
            USING (
                SELECT r.ClientId,
                       CAST(DATEADD(DAY, -n.n, GETDATE()) AS DATE) AS UsageDate,
                       50 AS SmsCount
                FROM ranked r
                CROSS JOIN (VALUES(1),(2),(3),(4),(5),(6),(7)) n(n)
                WHERE r.slot < 11
            ) src ON tgt.ClientId = src.ClientId AND tgt.UsageDate = src.UsageDate
            WHEN MATCHED     THEN UPDATE SET SmsCount = src.SmsCount
            WHEN NOT MATCHED THEN INSERT (ClientId, UsageDate, SmsCount, CreatedDate)
                                  VALUES (src.ClientId, src.UsageDate, src.SmsCount, GETDATE());
        """)

        # ── 1. Spike clients: today = 500 — slots 0-5 ───────────────────────
        cursor.execute(_CTE + """
            MERGE ClientSmsUsageDaily_Vtb AS tgt
            USING (
                SELECT ClientId, CAST(GETDATE() AS DATE) AS UsageDate, 500 AS SmsCount
                FROM ranked WHERE slot < 6
            ) src ON tgt.ClientId = src.ClientId AND tgt.UsageDate = src.UsageDate
            WHEN MATCHED     THEN UPDATE SET SmsCount = src.SmsCount
            WHEN NOT MATCHED THEN INSERT (ClientId, UsageDate, SmsCount, CreatedDate)
                                  VALUES (src.ClientId, src.UsageDate, src.SmsCount, GETDATE());
        """)
        spike_n = cursor.rowcount

        # ── 2. Drop clients: today = 3 — slots 6-10 ─────────────────────────
        cursor.execute(_CTE + """
            MERGE ClientSmsUsageDaily_Vtb AS tgt
            USING (
                SELECT ClientId, CAST(GETDATE() AS DATE) AS UsageDate, 3 AS SmsCount
                FROM ranked WHERE slot >= 6 AND slot < 11
            ) src ON tgt.ClientId = src.ClientId AND tgt.UsageDate = src.UsageDate
            WHEN MATCHED     THEN UPDATE SET SmsCount = src.SmsCount
            WHEN NOT MATCHED THEN INSERT (ClientId, UsageDate, SmsCount, CreatedDate)
                                  VALUES (src.ClientId, src.UsageDate, src.SmsCount, GETDATE());
        """)
        drop_n = cursor.rowcount

        # ── 3. Drop clients: yesterday = 250 — slots 6-10 ───────────────────
        #    (ensures yest > 50 AND avg > 20 conditions for drop detection)
        cursor.execute(_CTE + """
            MERGE ClientSmsUsageDaily_Vtb AS tgt
            USING (
                SELECT ClientId,
                       CAST(DATEADD(DAY, -1, GETDATE()) AS DATE) AS UsageDate,
                       250 AS SmsCount
                FROM ranked WHERE slot >= 6 AND slot < 11
            ) src ON tgt.ClientId = src.ClientId AND tgt.UsageDate = src.UsageDate
            WHEN MATCHED     THEN UPDATE SET SmsCount = src.SmsCount
            WHEN NOT MATCHED THEN INSERT (ClientId, UsageDate, SmsCount, CreatedDate)
                                  VALUES (src.ClientId,
                                          CAST(DATEADD(DAY, -1, GETDATE()) AS DATE),
                                          src.SmsCount, GETDATE());
        """)

        # ── 4. Recalibrate ClientSmsCredit_Vtb to match the refreshed patterns ──
        # Recomputes AllocatedSMS per slot so UsagePercent stays in its intended
        # band (healthy / warning / critical) even after sync_sms_usage() has
        # overwritten the daily rows with real Smslog_Vtb counts.
        # Uses the same slot formula as seed_test_data() so the distribution is
        # always consistent.
        cursor.execute(_CTE + """
            , used7 AS (
                SELECT ClientId, ISNULL(SUM(SmsCount), 0) AS TotalUsed
                FROM ClientSmsUsageDaily_Vtb WITH (NOLOCK)
                WHERE UsageDate >= CAST(DATEADD(DAY, -6, GETDATE()) AS DATE)
                GROUP BY ClientId
            ),
            targets AS (
                SELECT c.ClientId,
                       c.ManualAlloc,
                       c.AllocatedSMS        AS existing_alloc,
                       r.slot,
                       ISNULL(u.TotalUsed, 0) AS TotalUsed,
                       CASE
                           WHEN r.slot < 6
                               THEN CASE WHEN ISNULL(u.TotalUsed, 0) > 0
                                         THEN u.TotalUsed * 7   ELSE 5000 END
                           WHEN r.slot < 11
                               THEN CASE WHEN ISNULL(u.TotalUsed, 0) > 0
                                         THEN u.TotalUsed * 5   ELSE 3000 END
                           WHEN r.slot = 11
                               THEN CASE WHEN ISNULL(u.TotalUsed, 0) > 0
                                         THEN u.TotalUsed * 3   ELSE 2000 END
                           WHEN r.slot = 12
                               THEN CASE WHEN ISNULL(u.TotalUsed, 0) > 0
                                         THEN CEILING(CAST(u.TotalUsed AS FLOAT) / 0.65)
                                         ELSE 1000 END
                           ELSE  CASE WHEN ISNULL(u.TotalUsed, 0) > 0
                                      THEN CEILING(CAST(u.TotalUsed AS FLOAT) / 0.95)
                                      ELSE 500 END
                       END AS new_alloc
                FROM ClientSmsCredit_Vtb c WITH (NOLOCK)
                INNER JOIN ranked r ON r.ClientId = c.ClientId
                LEFT  JOIN used7  u ON u.ClientId = c.ClientId
            )
            UPDATE c
            SET c.AllocatedSMS = CASE WHEN t.ManualAlloc = 0
                                      THEN t.new_alloc
                                      ELSE t.existing_alloc END,
                c.UsedSMS      = t.TotalUsed,
                c.RemainingSMS = CASE
                                     WHEN CASE WHEN t.ManualAlloc = 0
                                               THEN t.new_alloc
                                               ELSE t.existing_alloc END > t.TotalUsed
                                     THEN CASE WHEN t.ManualAlloc = 0
                                               THEN t.new_alloc
                                               ELSE t.existing_alloc END - t.TotalUsed
                                     ELSE 0 END,
                c.ModifiedDate = GETDATE()
            FROM ClientSmsCredit_Vtb c
            INNER JOIN targets t ON t.ClientId = c.ClientId
            WHERE t.slot < 11 OR t.ManualAlloc = 0
        """)

        conn.commit()
        cursor.close()
        conn.close()

        invalidate_dashboard_cache()
        print(f"[HOURLY] Patterns applied — {spike_n} spike, {drop_n} drop.")
        return {"refreshed": True, "spike": spike_n, "drop": drop_n}

    except Exception as e:
        import traceback
        print("[ERROR] refresh_hourly_test_usage failed:")
        traceback.print_exc()
        return None


def reset_daily_alert_flags():
    """
    Runs daily at midnight for all clients.
    - Spike/drop flags: reset every day (intra-day events).
    - Threshold flags (50/90): reset only if usage has since dropped below the threshold,
      so clients re-trigger an alert if they climb again after getting new credits.
    """
    try:
        conn   = get_connection()
        cursor = conn.cursor()

        t = get_thresholds()

        cursor.execute("""
            UPDATE ClientSmsCredit_Vtb
            SET AlertSpikeSent = 0, AlertDropSent = 0
        """)

        cursor.execute(
            "UPDATE ClientSmsCredit_Vtb SET Alert50Sent = 0 WHERE UsagePercent < ?",
            t["warning"]
        )

        cursor.execute(
            "UPDATE ClientSmsCredit_Vtb SET Alert90Sent = 0 WHERE UsagePercent < ?",
            t["critical"]
        )

        conn.commit()
        cursor.close()
        conn.close()
        invalidate_dashboard_cache()
        print("[RESET] Daily alert flags reset for all clients.")
        return True

    except Exception as e:
        print(f"[ERROR] reset_daily_alert_flags: {e}")
        return False


# ══════════════════════════════════════════════════════════════════════════════
#  USER MANAGEMENT
# ══════════════════════════════════════════════════════════════════════════════

def fetch_users() -> list:
    """Return all user rows for the admin user-management panel."""
    try:
        conn   = get_connection()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT Username, Role, IsActive, CreatedDate, LastLogin
            FROM SmsMonitorUsers_Vtb
            ORDER BY CreatedDate
        """)
        rows = cursor.fetchall()
        cursor.close()
        conn.close()
        return [
            {
                "username":   row[0],
                "role":       row[1],
                "is_active":  bool(row[2]),
                "created":    row[3].strftime("%d %b %Y") if row[3] else "—",
                "last_login": row[4].strftime("%d %b %Y %H:%M") if row[4] else "Never"
            }
            for row in rows
        ]
    except Exception as e:
        print(f"[DB ERROR] fetch_users: {e}")
        return []


def create_user(username: str, password: str, role: str):
    """
    Create a new user. Returns dict on success or {"error": "..."} on conflict.
    Returns None on DB failure.
    """
    try:
        from passlib.context import CryptContext
        _ctx   = CryptContext(schemes=["bcrypt"], deprecated="auto")
        hashed = _ctx.hash(password)

        conn   = get_connection()
        cursor = conn.cursor()

        cursor.execute(
            "SELECT 1 FROM SmsMonitorUsers_Vtb WHERE Username = ?", username
        )
        if cursor.fetchone():
            cursor.close()
            conn.close()
            return {"error": "Username already exists."}

        cursor.execute("""
            INSERT INTO SmsMonitorUsers_Vtb (Username, PasswordHash, Role)
            VALUES (?, ?, ?)
        """, username, hashed, role)
        conn.commit()
        cursor.close()
        conn.close()
        return {"username": username, "role": role, "is_active": True,
                "created": "Just now", "last_login": "Never"}
    except Exception as e:
        print(f"[DB ERROR] create_user: {e}")
        return None


def set_user_status(username: str, is_active: bool) -> dict:
    """
    Activate or deactivate a user.
    Returns {"ok": True} or {"error": "reason"}.
    """
    try:
        conn   = get_connection()
        cursor = conn.cursor()

        if not is_active:
            # Prevent removing the last active admin
            cursor.execute("""
                SELECT COUNT(*) FROM SmsMonitorUsers_Vtb
                WHERE Role = 'admin' AND IsActive = 1 AND Username != ?
            """, username)
            if cursor.fetchone()[0] == 0:
                cursor.close()
                conn.close()
                return {"error": "Cannot deactivate the only active admin account."}

        cursor.execute("""
            UPDATE SmsMonitorUsers_Vtb
            SET IsActive     = ?,
                TokenVersion = TokenVersion + 1
            WHERE Username = ?
        """, 1 if is_active else 0, username)
        conn.commit()
        cursor.close()
        conn.close()
        return {"ok": True}
    except Exception as e:
        print(f"[DB ERROR] set_user_status: {e}")
        return {"error": str(e)}


def reset_user_password(username: str, new_password: str) -> bool:
    """Hash and save a new password for the given user."""
    try:
        from passlib.context import CryptContext
        _ctx   = CryptContext(schemes=["bcrypt"], deprecated="auto")
        hashed = _ctx.hash(new_password)

        conn   = get_connection()
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE SmsMonitorUsers_Vtb
            SET PasswordHash = ?,
                TokenVersion = TokenVersion + 1
            WHERE Username = ?
        """, hashed, username)
        conn.commit()
        cursor.close()
        conn.close()
        return True
    except Exception as e:
        print(f"[DB ERROR] reset_user_password: {e}")
        return False


def bump_token_version(username: str) -> bool:
    """Increment TokenVersion for a user, immediately invalidating all their JWTs."""
    try:
        conn   = get_connection()
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE SmsMonitorUsers_Vtb
            SET TokenVersion = TokenVersion + 1
            WHERE Username = ?
        """, username)
        conn.commit()
        cursor.close()
        conn.close()
        return True
    except Exception as e:
        print(f"[DB ERROR] bump_token_version: {e}")
        return False


def fetch_audit_log(limit: int = 50, offset: int = 0,
                    user_filter: str = None, action_filter: str = None) -> dict:
    """
    Returns a paginated slice of AuditLog_Vtb, newest first.
    user_filter   — partial match on Username (LIKE %value%)
    action_filter — exact match on Action
    Returns {"total": int, "rows": [...]}
    """
    try:
        conn   = get_connection()
        cursor = conn.cursor()

        conditions = []
        params     = []

        if user_filter:
            conditions.append("Username LIKE ?")
            params.append(f"%{user_filter}%")

        if action_filter:
            conditions.append("Action = ?")
            params.append(action_filter)

        where = ("WHERE " + " AND ".join(conditions)) if conditions else ""

        cursor.execute(f"SELECT COUNT(*) FROM AuditLog_Vtb {where}", params)
        total = cursor.fetchone()[0]

        cursor.execute(f"""
            SELECT LogId, Username, Action, TargetId, Detail, IpAddress, CreatedDate
            FROM AuditLog_Vtb
            {where}
            ORDER BY CreatedDate DESC
            OFFSET ? ROWS FETCH NEXT ? ROWS ONLY
        """, params + [offset, limit])
        rows = cursor.fetchall()
        cursor.close()
        conn.close()

        return {
            "total": total,
            "rows": [
                {
                    "log_id":    row[0],
                    "username":  row[1],
                    "action":    row[2],
                    "target_id": row[3] or "",
                    "detail":    row[4] or "",
                    "ip":        row[5] or "",
                    "created":   row[6].strftime("%d %b %Y %H:%M:%S") if row[6] else ""
                }
                for row in rows
            ]
        }
    except Exception as e:
        print(f"[DB ERROR] fetch_audit_log: {e}")
        return {"total": 0, "rows": []}


def fetch_sms_trend(period: str) -> dict:
    """
    Returns grand-total SMS counts aggregated across all clients.
    period: 'weekly' (last 8 weeks), 'monthly' (last 12 months), 'yearly' (all years)
    Sourced from Smslog_Vtb for full historical coverage.
    Returns {"labels": [...], "data": [...], "total": int, "period": str}
    """
    MONTHS = ["Jan","Feb","Mar","Apr","May","Jun",
              "Jul","Aug","Sep","Oct","Nov","Dec"]
    try:
        conn   = get_connection()
        cursor = conn.cursor()

        if period == "weekly":
            cursor.execute("""
                SELECT
                    YEAR(s.SentTime)               AS Yr,
                    DATEPART(ISO_WEEK, s.SentTime) AS Wk,
                    MIN(CAST(s.SentTime AS DATE))  AS WeekStart,
                    SUM(CEILING(LEN(LTRIM(RTRIM(ISNULL(s.Message,'')))) / 160.0)) AS TotalSms
                FROM Smslog_Vtb s WITH (NOLOCK)
                WHERE s.SentTime >= DATEADD(WEEK, -12, GETDATE())
                GROUP BY YEAR(s.SentTime), DATEPART(ISO_WEEK, s.SentTime)
                ORDER BY Yr, Wk
            """)
            rows   = cursor.fetchall()
            labels = [r[2].strftime("%d %b") for r in rows]
            data   = [int(r[3]) for r in rows]

        elif period == "monthly":
            cursor.execute("""
                SELECT
                    YEAR(s.SentTime)  AS Yr,
                    MONTH(s.SentTime) AS Mo,
                    SUM(CEILING(LEN(LTRIM(RTRIM(ISNULL(s.Message,'')))) / 160.0)) AS TotalSms
                FROM Smslog_Vtb s WITH (NOLOCK)
                WHERE s.SentTime >= DATEADD(MONTH, -12, GETDATE())
                GROUP BY YEAR(s.SentTime), MONTH(s.SentTime)
                ORDER BY Yr, Mo
            """)
            rows   = cursor.fetchall()
            labels = [f"{MONTHS[r[1]-1]} {r[0]}" for r in rows]
            data   = [int(r[2]) for r in rows]

        elif period == "yearly":
            cursor.execute("""
                SELECT
                    YEAR(s.SentTime) AS Yr,
                    SUM(CEILING(LEN(LTRIM(RTRIM(ISNULL(s.Message,'')))) / 160.0)) AS TotalSms
                FROM Smslog_Vtb s WITH (NOLOCK)
                GROUP BY YEAR(s.SentTime)
                ORDER BY Yr
            """)
            rows   = cursor.fetchall()
            labels = [str(r[0]) for r in rows]
            data   = [int(r[1]) for r in rows]

        else:
            cursor.close()
            conn.close()
            return {"labels": [], "data": [], "total": 0, "period": period}

        cursor.close()
        conn.close()
        return {
            "labels": labels,
            "data":   data,
            "total":  sum(data),
            "period": period,
        }

    except Exception as e:
        print(f"[DB ERROR] fetch_sms_trend: {e}")
        return {"labels": [], "data": [], "total": 0, "period": period}
