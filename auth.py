"""
Authentication and user management for CV Screener.
Uses PostgreSQL for user storage, bcrypt passwords, JWT in httpOnly cookies.
DATABASE_URL is injected automatically by Render when a Postgres DB is attached.
"""
import os, secrets
from datetime import datetime, timedelta, timezone
from typing import Optional
from contextlib import contextmanager

import psycopg2
import psycopg2.extras
from fastapi import Cookie, HTTPException
from jose import JWTError, jwt
from passlib.context import CryptContext

# ── Config ─────────────────────────────────────────────────────────────────────
DATABASE_URL = os.environ.get("DATABASE_URL", "")
SECRET_KEY   = os.environ.get("JWT_SECRET", secrets.token_hex(32))
ALGORITHM    = "HS256"
TOKEN_TTL    = int(os.environ.get("TOKEN_TTL_HOURS", "8"))
COOKIE       = "qs_token"

pwd_ctx = CryptContext(schemes=["bcrypt"], deprecated="auto")

# ── Database ───────────────────────────────────────────────────────────────────
@contextmanager
def _conn():
    if not DATABASE_URL:
        raise RuntimeError("DATABASE_URL environment variable is not set.")
    # Render injects postgres:// but psycopg2 needs postgresql://
    url = DATABASE_URL.replace("postgres://", "postgresql://", 1)
    conn = psycopg2.connect(url, cursor_factory=psycopg2.extras.RealDictCursor)
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
        cur = conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id         SERIAL PRIMARY KEY,
                username   TEXT UNIQUE NOT NULL,
                email      TEXT UNIQUE NOT NULL,
                pw_hash    TEXT NOT NULL,
                role       TEXT NOT NULL DEFAULT 'recruiter',
                active     BOOLEAN NOT NULL DEFAULT TRUE,
                created_at TEXT NOT NULL
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS calibration (
                id         SERIAL PRIMARY KEY,
                note       TEXT NOT NULL,
                created_by TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS screening_examples (
                id              SERIAL PRIMARY KEY,
                role_title      TEXT NOT NULL,
                candidate_name  TEXT NOT NULL,
                ai_score        INTEGER NOT NULL,
                final_decision  TEXT NOT NULL,
                summary         TEXT,
                strengths       TEXT,
                gaps            TEXT,
                recruiter_note  TEXT,
                created_by      TEXT NOT NULL,
                created_at      TEXT NOT NULL
            )
        """)
        cur.execute("""
            CREATE INDEX IF NOT EXISTS idx_examples_role
            ON screening_examples (lower(role_title))
        """)
        # Seed default admin on first run
        cur.execute("SELECT id FROM users WHERE username='admin'")
        if not cur.fetchone():
            cur.execute("""
                INSERT INTO users (username, email, pw_hash, role, active, created_at)
                VALUES (%s, %s, %s, 'admin', TRUE, %s)
            """, ("admin", "admin@questalliance.net",
                  pwd_ctx.hash("changeme123"),
                  datetime.now(timezone.utc).isoformat()))

# ── Password helpers ───────────────────────────────────────────────────────────
def verify_password(plain: str, hashed: str) -> bool:
    return pwd_ctx.verify(plain, hashed)

def hash_password(plain: str) -> str:
    return pwd_ctx.hash(plain)

# ── JWT helpers ────────────────────────────────────────────────────────────────
def create_token(user_id: int, username: str, role: str) -> str:
    exp = datetime.now(timezone.utc) + timedelta(hours=TOKEN_TTL)
    return jwt.encode(
        {"sub": str(user_id), "username": username, "role": role, "exp": exp},
        SECRET_KEY, algorithm=ALGORITHM,
    )

def decode_token(token: str) -> Optional[dict]:
    try:
        return jwt.decode(token, SECRET_KEY, algorithms=[ALGORITHM])
    except JWTError:
        return None

# ── Current-user dependency ────────────────────────────────────────────────────
def get_current_user(qs_token: Optional[str] = Cookie(default=None)) -> dict:
    if not qs_token:
        raise HTTPException(status_code=401, detail="Not authenticated")
    payload = decode_token(qs_token)
    if not payload:
        raise HTTPException(status_code=401, detail="Invalid or expired session")
    with _conn() as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT id, username, email, role, active FROM users WHERE id=%s",
            (int(payload["sub"]),)
        )
        row = cur.fetchone()
    if not row or not row["active"]:
        raise HTTPException(status_code=401, detail="Account disabled")
    return dict(row)

def require_admin(qs_token: Optional[str] = Cookie(default=None)) -> dict:
    u = get_current_user(qs_token)
    if u["role"] != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    return u

# ── User CRUD ──────────────────────────────────────────────────────────────────
def list_users() -> list:
    with _conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT id, username, email, role, active, created_at FROM users ORDER BY id")
        return [dict(r) for r in cur.fetchall()]

def create_user(username: str, email: str, password: str, role: str) -> dict:
    if role not in ("admin", "recruiter"):
        raise HTTPException(status_code=400, detail="Role must be admin or recruiter")
    try:
        with _conn() as conn:
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO users (username, email, pw_hash, role, active, created_at)
                VALUES (%s, %s, %s, %s, TRUE, %s) RETURNING id
            """, (username.strip(), email.strip().lower(),
                  hash_password(password), role,
                  datetime.now(timezone.utc).isoformat()))
            uid = cur.fetchone()["id"]
    except psycopg2.errors.UniqueViolation:
        raise HTTPException(status_code=409, detail="Username or email already exists")
    return {"id": uid, "username": username, "email": email, "role": role, "active": True}

def update_user(uid: int, role: Optional[str] = None, active: Optional[bool] = None,
                password: Optional[str] = None) -> dict:
    with _conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT id FROM users WHERE id=%s", (uid,))
        if not cur.fetchone():
            raise HTTPException(status_code=404, detail="User not found")
        if role is not None:
            if role not in ("admin", "recruiter"):
                raise HTTPException(status_code=400, detail="Invalid role")
            cur.execute("UPDATE users SET role=%s WHERE id=%s", (role, uid))
        if active is not None:
            cur.execute("UPDATE users SET active=%s WHERE id=%s", (active, uid))
        if password:
            cur.execute("UPDATE users SET pw_hash=%s WHERE id=%s", (hash_password(password), uid))
        cur.execute(
            "SELECT id, username, email, role, active, created_at FROM users WHERE id=%s", (uid,)
        )
        return dict(cur.fetchone())

def delete_user(uid: int, requesting_uid: int):
    if uid == requesting_uid:
        raise HTTPException(status_code=400, detail="Cannot delete your own account")
    with _conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT id FROM users WHERE id=%s", (uid,))
        if not cur.fetchone():
            raise HTTPException(status_code=404, detail="User not found")
        cur.execute("DELETE FROM users WHERE id=%s", (uid,))

def get_user_by_credentials(username: str, password: str) -> Optional[dict]:
    with _conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT * FROM users WHERE username=%s AND active=TRUE", (username,))
        row = cur.fetchone()
    if not row or not verify_password(password, row["pw_hash"]):
        return None
    return dict(row)

# ── Calibration notes ──────────────────────────────────────────────────────────
def get_calibration_notes() -> list:
    with _conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT id, note, created_by, created_at FROM calibration ORDER BY id")
        return [dict(r) for r in cur.fetchall()]

def add_calibration_note(note: str, username: str) -> dict:
    with _conn() as conn:
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO calibration (note, created_by, created_at) VALUES (%s, %s, %s) RETURNING id",
            (note.strip(), username, datetime.now(timezone.utc).isoformat())
        )
        nid = cur.fetchone()["id"]
    return {"id": nid, "note": note.strip(), "created_by": username}

def delete_calibration_note(note_id: int):
    with _conn() as conn:
        cur = conn.cursor()
        cur.execute("DELETE FROM calibration WHERE id=%s", (note_id,))

# ── Screening examples (JD-specific learning) ──────────────────────────────────

def init_examples_table(cur):
    cur.execute("""
        CREATE TABLE IF NOT EXISTS screening_examples (
            id              SERIAL PRIMARY KEY,
            role_title      TEXT NOT NULL,
            candidate_name  TEXT NOT NULL,
            ai_score        INTEGER NOT NULL,
            final_decision  TEXT NOT NULL,
            summary         TEXT,
            strengths       TEXT,
            gaps            TEXT,
            recruiter_note  TEXT,
            created_by      TEXT NOT NULL,
            created_at      TEXT NOT NULL
        )
    """)
    cur.execute("""
        CREATE INDEX IF NOT EXISTS idx_examples_role
        ON screening_examples (lower(role_title))
    """)

def save_screening_examples(role_title: str, examples: list, username: str):
    """examples: list of dicts with candidate_name, ai_score, final_decision,
       summary, strengths, gaps, recruiter_note."""
    now = datetime.now(timezone.utc).isoformat()
    with _conn() as conn:
        cur = conn.cursor()
        init_examples_table(cur)
        for ex in examples:
            cur.execute("""
                INSERT INTO screening_examples
                  (role_title, candidate_name, ai_score, final_decision,
                   summary, strengths, gaps, recruiter_note, created_by, created_at)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
            """, (
                role_title,
                ex.get("candidate_name", ""),
                int(ex.get("ai_score", 0)),
                ex.get("final_decision", ""),
                ex.get("summary", ""),
                ex.get("strengths", ""),
                ex.get("gaps", ""),
                ex.get("recruiter_note", ""),
                username,
                now,
            ))

def get_screening_examples(role_title: str, limit: int = 20) -> list:
    """Return the most recent saved examples for this role title."""
    with _conn() as conn:
        cur = conn.cursor()
        try:
            init_examples_table(cur)
            cur.execute("""
                SELECT candidate_name, ai_score, final_decision, summary,
                       strengths, gaps, recruiter_note, created_at
                FROM screening_examples
                WHERE lower(role_title) = lower(%s)
                ORDER BY id DESC LIMIT %s
            """, (role_title, limit))
            return [dict(r) for r in cur.fetchall()]
        except Exception:
            return []

def list_example_roles() -> list:
    """Return distinct role titles that have saved examples, with counts."""
    with _conn() as conn:
        cur = conn.cursor()
        try:
            init_examples_table(cur)
            cur.execute("""
                SELECT role_title, COUNT(*) as count, MAX(created_at) as last_saved
                FROM screening_examples
                GROUP BY role_title
                ORDER BY last_saved DESC
            """)
            return [dict(r) for r in cur.fetchall()]
        except Exception:
            return []
