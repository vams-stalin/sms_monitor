import os
from datetime import datetime, timedelta, timezone
from jose import JWTError, jwt
from passlib.context import CryptContext
from fastapi import Request, HTTPException

SECRET_KEY         = os.getenv("JWT_SECRET", "sms-monitor-change-this-in-production")
ALGORITHM          = "HS256"
TOKEN_EXPIRE_HOURS = 8

pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")


# ── Password helpers ──────────────────────────────────────────────────────────

def hash_password(plain: str) -> str:
    return pwd_context.hash(plain)


def verify_password(plain: str, hashed: str) -> bool:
    return pwd_context.verify(plain, hashed)


# ── JWT helpers ───────────────────────────────────────────────────────────────

def create_token(username: str, role: str, token_version: int) -> str:
    payload = {
        "sub":  username,
        "role": role,
        "ver":  token_version,
        "exp":  datetime.now(timezone.utc) + timedelta(hours=TOKEN_EXPIRE_HOURS)
    }
    return jwt.encode(payload, SECRET_KEY, algorithm=ALGORITHM)


def decode_token(token: str) -> dict | None:
    try:
        return jwt.decode(token, SECRET_KEY, algorithms=[ALGORITHM])
    except JWTError:
        return None


def get_current_user(request: Request) -> dict | None:
    """
    Decodes the JWT from the request cookie, then verifies the token version
    against the DB. Returns the payload dict or None if invalid/revoked.
    A revoked token (deactivated user, password reset, or logout) returns None
    immediately regardless of the JWT expiry time.
    """
    token = request.cookies.get("access_token")
    if not token:
        return None
    payload = decode_token(token)
    if not payload:
        return None

    # Lazy import avoids a circular dependency (database.py never imports auth.py)
    from database import get_user_by_username
    user = get_user_by_username(payload.get("sub", ""))
    if not user or not user["is_active"]:
        return None
    if payload.get("ver") != user["token_version"]:
        return None

    return payload


# ── FastAPI dependencies ──────────────────────────────────────────────────────

def require_auth_api(request: Request) -> dict:
    """Dependency for API routes — raises 401 if not authenticated or revoked."""
    user = get_current_user(request)
    if not user:
        raise HTTPException(status_code=401, detail="Not authenticated")
    return user


def require_admin_api(request: Request) -> dict:
    """Dependency for admin-only API routes — raises 403 if not admin."""
    user = get_current_user(request)
    if not user:
        raise HTTPException(status_code=401, detail="Not authenticated")
    if user.get("role") != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    return user
