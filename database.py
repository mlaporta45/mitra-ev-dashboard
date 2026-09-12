import sqlite3
import json
from pathlib import Path
from datetime import datetime
from typing import Optional

DB_PATH = Path(__file__).parent / "mitra_ev.db"

SCHEMA = """
CREATE TABLE IF NOT EXISTS documents (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    filename TEXT NOT NULL,
    uploaded_at TEXT NOT NULL,
    period TEXT,
    file_hash TEXT UNIQUE
);

CREATE TABLE IF NOT EXISTS metrics (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    document_id INTEGER REFERENCES documents(id),
    period TEXT NOT NULL,
    metric_name TEXT NOT NULL,
    value REAL,
    value_text TEXT,
    unit TEXT,
    extracted_at TEXT NOT NULL,
    UNIQUE(period, metric_name)
);

CREATE TABLE IF NOT EXISTS covenants (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    document_id INTEGER REFERENCES documents(id),
    period TEXT NOT NULL,
    covenant_name TEXT NOT NULL,
    actual_value REAL,
    threshold REAL,
    threshold_text TEXT,
    status TEXT,
    extracted_at TEXT NOT NULL,
    UNIQUE(period, covenant_name)
);
"""

METRIC_NAMES = [
    "EV Vehicle Leases Revenue",
    "DCFC Charging Fees Revenue",
    "LCFS Credits Revenue",
    "Other Revenue",
    "Total Revenue",
    "EBITDA",
    "Cash Position",
    "Total Debt",
    "Vehicles in Service",
    "Vehicles Under MLA",
    "Truck MRR",
    "DCFC in Service",
    "DCFC Under SHA",
    "DCFC Avg. Utilization Rate",
    "Minimum Liquidity",
    "Tangible Net Worth",
    "DSCR",
]

COVENANT_DEFINITIONS = {
    "Minimum Liquidity": {"threshold": 1_000_000, "threshold_text": "$1mm", "direction": ">="},
    "Tangible Net Worth": {"threshold": None, "threshold_text": "TBD", "direction": ">="},
    "DSCR": {"threshold": 1.25, "threshold_text": "1.25x", "direction": ">="},
}


def get_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_connection()
    conn.executescript(SCHEMA)
    conn.commit()
    conn.close()


def file_already_processed(file_hash: str) -> bool:
    conn = get_connection()
    try:
        row = conn.execute("SELECT id FROM documents WHERE file_hash=?", (file_hash,)).fetchone()
        return row is not None
    finally:
        conn.close()


def insert_document(filename: str, period: str, file_hash: str) -> Optional[int]:
    conn = get_connection()
    try:
        cur = conn.execute(
            "INSERT OR IGNORE INTO documents (filename, uploaded_at, period, file_hash) VALUES (?, ?, ?, ?)",
            (filename, datetime.utcnow().isoformat(), period, file_hash),
        )
        conn.commit()
        if cur.lastrowid:
            return cur.lastrowid
        row = conn.execute("SELECT id FROM documents WHERE file_hash=?", (file_hash,)).fetchone()
        return row["id"] if row else None
    finally:
        conn.close()


def upsert_metrics(document_id: int, period: str, metrics: dict):
    now = datetime.utcnow().isoformat()
    conn = get_connection()
    try:
        for name, info in metrics.items():
            if isinstance(info, dict):
                value = info.get("value")
                value_text = info.get("value_text") or str(value) if value is not None else info.get("raw")
                unit = info.get("unit", "")
            else:
                value = info if isinstance(info, (int, float)) else None
                value_text = str(info)
                unit = ""
            conn.execute(
                """INSERT INTO metrics (document_id, period, metric_name, value, value_text, unit, extracted_at)
                   VALUES (?, ?, ?, ?, ?, ?, ?)
                   ON CONFLICT(period, metric_name) DO UPDATE SET
                     value=excluded.value,
                     value_text=excluded.value_text,
                     unit=excluded.unit,
                     document_id=excluded.document_id,
                     extracted_at=excluded.extracted_at""",
                (document_id, period, name, value, value_text, unit, now),
            )
        conn.commit()
    finally:
        conn.close()


def upsert_covenants(document_id: int, period: str, covenants: dict):
    now = datetime.utcnow().isoformat()
    conn = get_connection()
    try:
        for name, info in covenants.items():
            if not isinstance(info, dict):
                continue
            actual = info.get("actual")
            defn = COVENANT_DEFINITIONS.get(name, {})
            threshold = info.get("threshold") or defn.get("threshold")
            threshold_text = info.get("threshold_text") or defn.get("threshold_text", "")
            direction = defn.get("direction", ">=")
            status = info.get("status")
            if status is None and actual is not None and threshold is not None:
                if direction == ">=":
                    status = "pass" if actual >= threshold else "fail"
                else:
                    status = "pass" if actual <= threshold else "fail"
            conn.execute(
                """INSERT INTO covenants (document_id, period, covenant_name, actual_value, threshold, threshold_text, status, extracted_at)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                   ON CONFLICT(period, covenant_name) DO UPDATE SET
                     actual_value=excluded.actual_value,
                     threshold=excluded.threshold,
                     threshold_text=excluded.threshold_text,
                     status=excluded.status,
                     document_id=excluded.document_id,
                     extracted_at=excluded.extracted_at""",
                (document_id, period, name, actual, threshold, threshold_text, status, now),
            )
        conn.commit()
    finally:
        conn.close()


def get_all_periods() -> list[str]:
    conn = get_connection()
    try:
        rows = conn.execute(
            "SELECT DISTINCT period FROM metrics ORDER BY period"
        ).fetchall()
        return [r["period"] for r in rows]
    finally:
        conn.close()


def get_metrics_for_period(period: str) -> dict:
    conn = get_connection()
    try:
        rows = conn.execute(
            "SELECT metric_name, value, value_text, unit FROM metrics WHERE period=?", (period,)
        ).fetchall()
        return {r["metric_name"]: {"value": r["value"], "value_text": r["value_text"], "unit": r["unit"]} for r in rows}
    finally:
        conn.close()


def get_covenants_for_period(period: str) -> dict:
    conn = get_connection()
    try:
        rows = conn.execute(
            "SELECT covenant_name, actual_value, threshold, threshold_text, status FROM covenants WHERE period=?",
            (period,),
        ).fetchall()
        return {
            r["covenant_name"]: {
                "actual": r["actual_value"],
                "threshold": r["threshold"],
                "threshold_text": r["threshold_text"],
                "status": r["status"],
            }
            for r in rows
        }
    finally:
        conn.close()


def get_metric_timeseries(metric_name: str) -> list[dict]:
    conn = get_connection()
    try:
        rows = conn.execute(
            "SELECT period, value, value_text, unit FROM metrics WHERE metric_name=? ORDER BY period",
            (metric_name,),
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def get_all_metrics_wide() -> "pd.DataFrame":
    import pandas as pd
    conn = get_connection()
    try:
        df = pd.read_sql_query(
            "SELECT period, metric_name, value, value_text, unit FROM metrics ORDER BY period, metric_name",
            conn,
        )
        return df
    finally:
        conn.close()


def get_documents() -> list[dict]:
    conn = get_connection()
    try:
        rows = conn.execute(
            "SELECT id, filename, uploaded_at, period FROM documents ORDER BY uploaded_at DESC"
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


init_db()
