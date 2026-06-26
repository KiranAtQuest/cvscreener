"""
Authentication and user management for CV Screener.
Uses SQLite for user storage, bcrypt for passwords, JWT in httpOnly cookies.
"""
import os, sqlite3, secrets
from datetime import datetime, timedelta, timezone
from typing import Optional

from fastapi import Cookie, HTTPException, Request, status
from jose import JWTError, jwt
from passlib.context import CryptContext

# ── Config ─────────────────────────────────────────────────────────────────────
DB_PATH    = os.environ.get("AUTH_DB_PATH", "users.db")
SECRET_KEY = os.environ.get("JWT_SECRET", secrets.token_hex(32))
ALGORITHM  = "HS256"
TOKEN_TTL  = int(os.environ.get("TOKEN_TTL_HOURS", "8"))
COOKIE     = "qs_token"

pwd_ctx = CryptContext(schemes=["bcrypt"], deprecated="auto")

# ── Database ───────────────────────────────────────────────────────────────────
def _conn():
    c = sqlite3.connect(DB_PATH)
    c.row_factory = sqlite3.Row
    return c

def init_db():
    with _conn() as c:
        c.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                username   TEXT UNIQUE NOT NULL,
                email      TEXT UNIQUE NOT NULL,
                pw_hash    TEXT NOT NULL,
                role       TEXT NOT NULL DEFAULT 'recruiter',
                active     INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL
            )
        """)
        c.execute("""
            CREATE TABLE IF NOT EXISTS calibration (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                note       TEXT NOT NULL,
                created_by TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
        """)
        # Seed default admin on first run
        row = c.execute("SELECT id FROM users WHERE username='admin'").fetchone()
        if not row:
            c.execute("""
                INSERT INTO users (username, email, pw_hash, role, active, created_at)
                VALUES (?, ?, ?, 'admin', 1, ?)
            """, ("admin", "admin@questalliance.net",
                  pwd_ctx.hash("changeme123"),
                  datetime.now(timezone.utc).isoformat()))
            c.commit()

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
    # Verify user still active in DB
    with _conn() as c:
        row = c.execute(
            "SELECT id, username, email, role, active FROM users WHERE id=?",
            (int(payload["sub"]),)
        ).fetchone()
    if not row or not row["active"]:
        raise HTTPException(status_code=401, detail="Account disabled")
    return dict(row)

def require_admin(user: dict = None, qs_token: Optional[str] = Cookie(default=None)) -> dict:
    u = get_current_user(qs_token)
    if u["role"] != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    return u

# ── User CRUD ──────────────────────────────────────────────────────────────────
def list_users() -> list:
    with _conn() as c:
        rows = c.execute(
            "SELECT id, username, email, role, active, created_at FROM users ORDER BY id"
        ).fetchall()
    return [dict(r) for r in rows]

def create_user(username: str, email: str, password: str, role: str) -> dict:
    if role not in ("admin", "recruiter"):
        raise HTTPException(status_code=400, detail="Role must be admin or recruiter")
    try:
        with _conn() as c:
            c.execute("""
                INSERT INTO users (username, email, pw_hash, role, active, created_at)
                VALUES (?, ?, ?, ?, 1, ?)
            """, (username.strip(), email.strip().lower(),
                  hash_password(password), role,
                  datetime.now(timezone.utc).isoformat()))
            c.commit()
            uid = c.execute("SELECT last_insert_rowid()").fetchone()[0]
    except sqlite3.IntegrityError as e:
        raise HTTPException(status_code=409, detail="Username or email already exists")
    return {"id": uid, "username": username, "email": email, "role": role, "active": True}

def update_user(uid: int, role: Optional[str] = None, active: Optional[bool] = None,
                password: Optional[str] = None) -> dict:
    with _conn() as c:
        row = c.execute("SELECT * FROM users WHERE id=?", (uid,)).fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="User not found")
        if role is not None:
            if role not in ("admin", "recruiter"):
                raise HTTPException(status_code=400, detail="Invalid role")
            c.execute("UPDATE users SET role=? WHERE id=?", (role, uid))
        if active is not None:
            c.execute("UPDATE users SET active=? WHERE id=?", (1 if active else 0, uid))
        if password:
            c.execute("UPDATE users SET pw_hash=? WHERE id=?", (hash_password(password), uid))
        c.commit()
        updated = c.execute(
            "SELECT id, username, email, role, active, created_at FROM users WHERE id=?", (uid,)
        ).fetchone()
    return dict(updated)

def delete_user(uid: int, requesting_uid: int):
    if uid == requesting_uid:
        raise HTTPException(status_code=400, detail="Cannot delete your own account")
    with _conn() as c:
        row = c.execute("SELECT id FROM users WHERE id=?", (uid,)).fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="User not found")
        c.execute("DELETE FROM users WHERE id=?", (uid,))
        c.commit()

def get_user_by_credentials(username: str, password: str) -> Optional[dict]:
    with _conn() as c:
        row = c.execute(
            "SELECT * FROM users WHERE username=? AND active=1", (username,)
        ).fetchone()
    if not row or not verify_password(password, row["pw_hash"]):
        return None
    return dict(row)

# ── Calibration notes ──────────────────────────────────────────────────────────

def get_calibration_notes() -> list:
    with _conn() as c:
        rows = c.execute(
            "SELECT id, note, created_by, created_at FROM calibration ORDER BY id"
        ).fetchall()
    return [dict(r) for r in rows]

def add_calibration_note(note: str, username: str) -> dict:
    with _conn() as c:
        c.execute(
            "INSERT INTO calibration (note, created_by, created_at) VALUES (?, ?, ?)",
            (note.strip(), username, datetime.now(timezone.utc).isoformat())
        )
        c.commit()
        nid = c.execute("SELECT last_insert_rowid()").fetchone()[0]
    return {"id": nid, "note": note.strip(), "created_by": username}

def delete_calibration_note(note_id: int):
    with _conn() as c:
        c.execute("DELETE FROM calibration WHERE id=?", (note_id,))
        c.commit()
