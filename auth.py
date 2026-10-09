import os, contextlib, secrets, smtplib
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from typing import Optional

import bcrypt
import psycopg2
from psycopg2.extras import RealDictCursor
from jose import jwt, JWTError
from fastapi import Cookie, HTTPException

DATABASE_URL    = os.environ.get("DATABASE_URL", "")
JWT_SECRET      = os.environ.get("JWT_SECRET", "dev-secret-change-me")
JWT_ALG         = "HS256"
JWT_EXPIRE_H    = 12
SMTP_HOST       = os.environ.get("SMTP_HOST", "")
SMTP_PORT       = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER       = os.environ.get("SMTP_USER", "")
SMTP_PASS       = os.environ.get("SMTP_PASS", "")
SMTP_FROM       = os.environ.get("SMTP_FROM", SMTP_USER)
OTP_EXPIRE_MIN  = 10


def _hash_pw(plain: str) -> str:
    return bcrypt.hashpw(plain.encode(), bcrypt.gensalt(12)).decode()

def _verify_pw(plain: str, hashed: str) -> bool:
    try:
        return bcrypt.checkpw(plain.encode(), hashed.encode())
    except Exception:
        return False


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
            cur.execute("""
                CREATE TABLE IF NOT EXISTS saved_jds (
                    id SERIAL PRIMARY KEY,
                    role_title TEXT NOT NULL DEFAULT '',
                    jd_text TEXT NOT NULL,
                    created_by TEXT,
                    created_at TIMESTAMPTZ DEFAULT NOW()
                )
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS login_otps (
                    id SERIAL PRIMARY KEY,
                    username TEXT NOT NULL,
                    otp_code TEXT NOT NULL,
                    expires_at TIMESTAMPTZ NOT NULL,
                    used BOOLEAN NOT NULL DEFAULT FALSE,
                    created_at TIMESTAMPTZ DEFAULT NOW()
                )
            """)
            # Ensure column defaults exist (may be missing on older deployments)
            cur.execute("ALTER TABLE users ALTER COLUMN created_at SET DEFAULT NOW()")
            cur.execute("ALTER TABLE users ALTER COLUMN active SET DEFAULT TRUE")
            cur.execute("UPDATE users SET active=TRUE WHERE active IS NULL")
            # Seed default admin
            cur.execute("SELECT id FROM users WHERE username = 'admin'")
            if not cur.fetchone():
                cur.execute(
                    "INSERT INTO users (username, email, pw_hash, role) VALUES (%s,%s,%s,%s)",
                    ("admin", "admin@questalliance.net", _hash_pw("changeme123"), "admin")
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
    return _verify_pw(plain, hashed)


def login_user(username: str, password: str):
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM users WHERE username=%s AND active=TRUE", (username,))
            user = cur.fetchone()
    if not user or not verify_password(password, user["pw_hash"]):
        return None
    return dict(user)


# ── OTP login ──────────────────────────────────────────────────────────────────

def get_user_by_username(username_or_email: str):
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT id, username, email, role, active FROM users WHERE (username=%s OR email=%s) AND active=TRUE",
                (username_or_email, username_or_email)
            )
            row = cur.fetchone()
    return dict(row) if row else None


def issue_otp(username: str) -> str:
    """Generate a 6-digit OTP, store it, and return the code (caller sends the email)."""
    code = f"{secrets.randbelow(900000) + 100000}"
    with _conn() as conn:
        with conn.cursor() as cur:
            # Invalidate previous unused OTPs for this user
            cur.execute("UPDATE login_otps SET used=TRUE WHERE username=%s AND used=FALSE", (username,))
            cur.execute(
                "INSERT INTO login_otps (username, otp_code, expires_at) VALUES (%s, %s, NOW() + INTERVAL '%s minutes')",
                (username, code, OTP_EXPIRE_MIN)
            )
    return code


def verify_otp(username: str, code: str) -> bool:
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT id FROM login_otps
                WHERE username=%s AND otp_code=%s AND used=FALSE AND expires_at > NOW()
            """, (username, code))
            row = cur.fetchone()
            if not row:
                return False
            cur.execute("UPDATE login_otps SET used=TRUE WHERE id=%s", (row["id"],))
    return True


def send_otp_email(to_email: str, username: str, code: str):
    if not SMTP_HOST or not SMTP_USER:
        raise RuntimeError("SMTP not configured — set SMTP_HOST, SMTP_USER, SMTP_PASS env vars on Render.")
    body = (
        f"Hi {username},\n\n"
        f"Your CV Screener login code is:\n\n"
        f"  {code}\n\n"
        f"This code expires in {OTP_EXPIRE_MIN} minutes. Do not share it.\n\n"
        f"— Quest Alliance CV Screener"
    )
    msg = MIMEText(body)
    msg["Subject"] = f"{code} — your CV Screener login code"
    msg["From"]    = SMTP_FROM
    msg["To"]      = to_email
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=15) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.ehlo()
        smtp.login(SMTP_USER, SMTP_PASS)
        smtp.send_message(msg)


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
                "INSERT INTO users (username, email, pw_hash, role, active) VALUES (%s,%s,%s,%s,TRUE) RETURNING id",
                (username, email, _hash_pw(password), role)
            )
            return cur.fetchone()["id"]


def update_user(uid: int, **fields):
    if "password" in fields:
        fields["pw_hash"] = _hash_pw(fields.pop("password"))
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


# ── Saved JDs ──────────────────────────────────────────────────────────────────

def list_jds():
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT id, role_title, jd_text, created_by, created_at FROM saved_jds ORDER BY created_at DESC LIMIT 50")
            return [dict(r) for r in cur.fetchall()]


def save_jd(role_title: str, jd_text: str, username: str):
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "INSERT INTO saved_jds (role_title, jd_text, created_by) VALUES (%s,%s,%s) RETURNING id",
                (role_title, jd_text, username)
            )
            return cur.fetchone()["id"]


def delete_jd(jd_id: int):
    with _conn() as conn:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM saved_jds WHERE id=%s", (jd_id,))


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
