import sqlite3
import os
from datetime import datetime

try:
    from app.paths import get_db_path
except ImportError:
    from paths import get_db_path

DB_PATH = get_db_path()

def get_connection():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    return sqlite3.connect(DB_PATH)

def init_db():
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS qc_records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp TEXT,
                image_name TEXT UNIQUE,
                image_path TEXT,
                status TEXT,
                operator_id TEXT,
                batch_id TEXT,
                reject_reason TEXT,
                reviewed_at TEXT
            )
        ''')
        conn.commit()

def insert_capture(image_name, image_path):
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO qc_records (timestamp, image_name, image_path, status)
            VALUES (?, ?, ?, ?)
        ''', (timestamp, image_name, image_path, 'WAITING'))
        conn.commit()
    return timestamp

def update_classification(image_name, status, operator_id, batch_id, reject_reason=None):
    reviewed_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute('''
            UPDATE qc_records
            SET status = ?, operator_id = ?, batch_id = ?, reject_reason = ?, reviewed_at = ?
            WHERE image_name = ?
        ''', (status, operator_id, batch_id, reject_reason, reviewed_at, image_name))
        conn.commit()
    return reviewed_at

def undo_classification(image_name):
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute('''
            UPDATE qc_records
            SET status = "WAITING", operator_id = NULL, batch_id = NULL, reject_reason = NULL, reviewed_at = NULL
            WHERE image_name = ?
        ''', (image_name,))
        conn.commit()

def delete_classification(image_name):
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute('''
            DELETE FROM qc_records
            WHERE image_name = ?
        ''', (image_name,))
        conn.commit()

def get_waiting_count():
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute('SELECT COUNT(*) FROM qc_records WHERE status = "WAITING"')
        return cursor.fetchone()[0]

def get_stats():
    """Returns stats scoped to TODAY's date only, so each shift starts fresh."""
    today = datetime.now().strftime('%Y-%m-%d')
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute('''
            SELECT
                SUM(CASE WHEN status = 'ACCEPTED' THEN 1 ELSE 0 END) as accepted,
                SUM(CASE WHEN status = 'REJECTED' THEN 1 ELSE 0 END) as rejected,
                SUM(CASE WHEN status = 'WAITING'  THEN 1 ELSE 0 END) as waiting,
                SUM(CASE WHEN status = 'REWORK'   THEN 1 ELSE 0 END) as rework,
                COUNT(*) as total
            FROM qc_records
            WHERE DATE(timestamp) = ?
        ''', (today,))
        row = cursor.fetchone()
        accepted = row[0] or 0
        rejected = row[1] or 0
        waiting  = row[2] or 0
        rework   = row[3] or 0
        total    = row[4] or 0
        total_reviewed = accepted + rejected + rework
        acceptance_rate = round((accepted / total_reviewed * 100), 2) if total_reviewed > 0 else 0.0

        # Most common reject reason today
        cursor.execute('''
            SELECT reject_reason, COUNT(*) as count
            FROM qc_records
            WHERE status="REJECTED" AND reject_reason IS NOT NULL AND reject_reason != ""
              AND DATE(timestamp) = ?
            GROUP BY reject_reason
            ORDER BY count DESC
            LIMIT 1
        ''', (today,))
        top_reason_row = cursor.fetchone()
        top_reject_reason = top_reason_row[0] if top_reason_row else "None"

        return {
            'total_captured': total,
            'accepted': accepted,
            'rejected': rejected,
            'rework':   rework,
            'waiting':  waiting,
            'rate': acceptance_rate,
            'top_reject_reason': top_reject_reason,
            'date': today
        }

def get_report_data():
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM qc_records ORDER BY timestamp ASC')
        columns = [desc[0] for desc in cursor.description]
        return [dict(zip(columns, row)) for row in cursor.fetchall()]

def get_batch_stats():
    """Returns batch stats scoped to TODAY's date only."""
    today = datetime.now().strftime('%Y-%m-%d')
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute('''
            SELECT 
                batch_id,
                SUM(CASE WHEN status = 'ACCEPTED' THEN 1 ELSE 0 END) as accepted,
                SUM(CASE WHEN status = 'REJECTED' THEN 1 ELSE 0 END) as rejected
            FROM qc_records
            WHERE batch_id IS NOT NULL AND batch_id != ""
              AND DATE(timestamp) = ?
            GROUP BY batch_id
        ''', (today,))
        return cursor.fetchall()
