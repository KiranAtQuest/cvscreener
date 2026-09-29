import os, contextlib
from datetime import datetime, timedelta
from typing import Optional

import psycopg2
from psycopg2.extras import RealDictCursor
from passlib.context import CryptContext
from jose import jwt, JWTError
from fastapi import Cookie, HTTPException

DATABASE_URL = os.environ.get("DATABASE_URL", "")
JWT_SECRET   = os.environ.get("JWT_SECRET", "dev-secret-change-me")
JWT_ALG      = "HS256"
JWT_EXPIRE_H = 12

pwd_ctx = CryptContext(schemes=["bcrypt"], deprecated="auto")


@contextlib.contextmanager
def _conn():
    conn = psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db():
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS users (
                    id SERIAL PRIMARY KEY,
                    username TEXT UNIQUE NOT NULL,
                    email TEXT,
                    pw_hash TEXT NOT NULL,
                    role TEXT NOT NULL DEFAULT 'recruiter',
                    active BOOLEAN NOT NULL DEFAULT TRUE,
                    created_at TIMESTAMPTZ DEFAULT NOW()
                )
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS calibration (
                    id SERIAL PRIMARY KEY,
                    note TEXT NOT NULL,
                    jd_hash TEXT,
                    created_by TEXT,
                    created_at TIMESTAMPTZ DEFAULT NOW()
                )
            """)
            cur.execute("ALTER TABLE calibration ADD COLUMN IF NOT EXISTS jd_hash TEXT")
            cur.execute("""
                CREATE TABLE IF NOT EXISTS screening_examples (
                    id SERIAL PRIMARY KEY,
                    role_title TEXT NOT NULL,
                    candidate_name TEXT,
                    ai_score INTEGER,
                    final_decision TEXT,
                    summary TEXT,
                    strengths TEXT,
                    gaps TEXT,
                    recruiter_note TEXT,
                    created_by TEXT,
                    created_at TIMESTAMPTZ DEFAULT NOW()
                )
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS upload_errors (
                    id SERIAL PRIMARY KEY,
                    timestamp TIMESTAMPTZ DEFAULT NOW(),
                    filename TEXT,
                    file_type TEXT,
                    role_title TEXT,
                    uploaded_by TEXT,
                    error_code TEXT,
                    error_plain TEXT,
                    fix_suggestion TEXT,
                    status TEXT DEFAULT 'needs-action'
                )
            """)
            # Seed default admin
            cur.execute("SELECT id FROM users WHERE username = 'admin'")
            if not cur.fetchone():
                cur.execute(
                    "INSERT INTO users (username, email, pw_hash, role) VALUES (%s,%s,%s,%s)",
                    ("admin", "admin@questalliance.net", pwd_ctx.hash("changeme123"), "admin")
                )


# ── Auth helpers ───────────────────────────────────────────────────────────────

def make_token(username: str, role: str) -> str:
    payload = {
        "sub": username,
        "role": role,
        "exp": datetime.utcnow() + timedelta(hours=JWT_EXPIRE_H),
    }
    return jwt.encode(payload, JWT_SECRET, algorithm=JWT_ALG)


def get_current_user(qs_token: Optional[str] = Cookie(None)):
    if not qs_token:
        raise HTTPException(status_code=401, detail="Not authenticated")
    try:
        payload = jwt.decode(qs_token, JWT_SECRET, algorithms=[JWT_ALG])
        username = payload.get("sub")
    except JWTError:
        raise HTTPException(status_code=401, detail="Invalid or expired session")
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT id, username, email, role, active FROM users WHERE username=%s", (username,))
            user = cur.fetchone()
    if not user or not user["active"]:
        raise HTTPException(status_code=401, detail="User not found or disabled")
    return dict(user)


def require_admin(qs_token: Optional[str] = Cookie(None)):
    user = get_current_user(qs_token)
    if user["role"] != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    return user


def verify_password(plain: str, hashed: str) -> bool:
    return pwd_ctx.verify(plain, hashed)


def login_user(username: str, password: str):
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM users WHERE username=%s AND active=TRUE", (username,))
            user = cur.fetchone()
    if not user or not verify_password(password, user["pw_hash"]):
        return None
    return dict(user)


# ── User management ────────────────────────────────────────────────────────────

def list_users():
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT id, username, email, role, active, created_at FROM users ORDER BY id")
            return [dict(r) for r in cur.fetchall()]


def create_user(username: str, email: str, password: str, role: str):
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "INSERT INTO users (username, email, pw_hash, role) VALUES (%s,%s,%s,%s) RETURNING id",
                (username, email, pwd_ctx.hash(password), role)
            )
            return cur.fetchone()["id"]


def update_user(uid: int, **fields):
    if "password" in fields:
        fields["pw_hash"] = pwd_ctx.hash(fields.pop("password"))
    sets = ", ".join(f"{k}=%s" for k in fields)
    vals = list(fields.values()) + [uid]
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute(f"UPDATE users SET {sets} WHERE id=%s", vals)


def delete_user(uid: int):
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM users WHERE id=%s AND username!='admin'", (uid,))


# ── Calibration ────────────────────────────────────────────────────────────────

def get_calibration_notes(jd_hash: str = None):
    with _conn() as conn:
        with conn.cursor() as cur:
            if jd_hash:
                cur.execute(
                    "SELECT id, note, created_by, created_at FROM calibration WHERE jd_hash=%s ORDER BY created_at",
                    (jd_hash,)
                )
            else:
                cur.execute("SELECT id, note, created_by, created_at FROM calibration WHERE jd_hash IS NULL ORDER BY created_at")
            return [dict(r) for r in cur.fetchall()]


def add_calibration_note(note: str, username: str, jd_hash: str = None):
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "INSERT INTO calibration (note, created_by, jd_hash) VALUES (%s,%s,%s) RETURNING id",
                (note, username, jd_hash)
            )
            return cur.fetchone()["id"]


def delete_calibration_note(note_id: int):
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM calibration WHERE id=%s", (note_id,))


# ── Screening examples (AI learning) ──────────────────────────────────────────

def save_screening_examples(role_title: str, examples: list, username: str):
    with _conn() as conn:
        with conn.cursor() as cur:
            for ex in examples:
                cur.execute("""
                    INSERT INTO screening_examples
                      (role_title, candidate_name, ai_score, final_decision, summary, strengths, gaps, recruiter_note, created_by)
                    VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)
                """, (
                    role_title,
                    ex.get("name", ""),
                    ex.get("ai_score"),
                    ex.get("final_decision", ""),
                    ex.get("summary", ""),
                    ex.get("strengths", ""),
                    ex.get("gaps", ""),
                    ex.get("recruiter_note", ""),
                    username,
                ))


def get_screening_examples(role_title: str, limit: int = 15):
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT candidate_name, ai_score, final_decision, summary, strengths, gaps, recruiter_note, created_at
                FROM screening_examples
                WHERE LOWER(role_title) = LOWER(%s)
                ORDER BY created_at DESC LIMIT %s
            """, (role_title, limit))
            return [dict(r) for r in cur.fetchall()]


# ── Upload error logging ───────────────────────────────────────────────────────

def log_upload_error(filename: str, file_type: str, role_title: str,
                     uploaded_by: str, error_code: str, error_plain: str, fix_suggestion: str):
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                INSERT INTO upload_errors
                  (filename, file_type, role_title, uploaded_by, error_code, error_plain, fix_suggestion)
                VALUES (%s,%s,%s,%s,%s,%s,%s)
            """, (filename, file_type, role_title, uploaded_by, error_code, error_plain, fix_suggestion))


def get_upload_errors(limit: int = 100):
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT id, timestamp, filename, file_type, role_title, uploaded_by,
                       error_plain, fix_suggestion, status
                FROM upload_errors ORDER BY timestamp DESC LIMIT %s
            """, (limit,))
            return [dict(r) for r in cur.fetchall()]


def update_error_status(error_id: int, status: str):
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute("UPDATE upload_errors SET status=%s WHERE id=%s", (status, error_id))
