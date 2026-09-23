import asyncio
import os
import re
from collections import defaultdict
from contextlib import asynccontextmanager
from datetime import datetime, timedelta
from pathlib import Path
from fastapi import FastAPI, Request, Form, Body, Depends
from fastapi.responses import HTMLResponse, RedirectResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

_BASE_DIR = Path(__file__).parent
from starlette.middleware.base import BaseHTTPMiddleware
from apscheduler.schedulers.background import BackgroundScheduler
from database import (fetch_clients, assign_sms_credits, fetch_dashboard_clients,
                      fetch_client_analytics, sync_sms_usage,
                      reset_daily_alert_flags, fetch_active_alerts,
                      notifier, ensure_schema, ensure_today_yesterday_data, get_last_synced,
                      get_user_by_username, update_last_login, log_audit,
                      fetch_users, create_user, set_user_status, reset_user_password,
                      bump_token_version, fetch_audit_log,
                      get_thresholds, set_thresholds,
                      invalidate_dashboard_cache, fetch_sms_trend,
                      expire_validity_credits,
                      get_client_report_email, save_client_report_email,
                      fetch_all_client_report_emails,
                      fetch_client_usage_report, build_usage_report_excel,
                      generate_detailed_sms_report)
from auth import (create_token, verify_password, get_current_user,
                  require_auth_api, require_admin_api)
from slowapi import Limiter, _rate_limit_exceeded_handler
from slowapi.util import get_remote_address
from slowapi.errors import RateLimitExceeded

# ── .env validation — warn about missing/insecure values at startup ──────────
def _validate_env():
    issues = []

    jwt = os.getenv("JWT_SECRET", "")
    if not jwt or jwt == "sms-monitor-change-this-in-production":
        issues.append("JWT_SECRET is not set or is the insecure default")

    smtp_host = os.getenv("SMTP_HOST", "")
    smtp_user = os.getenv("SMTP_USER", "")
    smtp_pass = os.getenv("SMTP_PASSWORD", "")
    if not smtp_host or not smtp_user or not smtp_pass:
        issues.append("SMTP_HOST / SMTP_USER / SMTP_PASSWORD are not fully configured — email alerts will not work")

    if not os.getenv("DB_SERVER", ""):
        issues.append("DB_SERVER is not set — database connection will fail")

    if not os.getenv("ALERT_TO", "").strip():
        issues.append("ALERT_TO is not set in email.env — alert emails will not be delivered")

    if os.getenv("DEFAULT_ADMIN_PASSWORD", "Admin@123") == "Admin@123":
        issues.append("DEFAULT_ADMIN_PASSWORD is still the factory default 'Admin@123' — set a strong password in .env")

    if issues:
        print("")
        print("[STARTUP] ================================================")
        print("[STARTUP]  ENV VALIDATION -- action required:")
        for i, msg in enumerate(issues, 1):
            print(f"[STARTUP]  {i}. {msg}")
        print("[STARTUP] ================================================")
        print("")

_validate_env()

# ── Login brute-force guard — in-memory, resets on server restart ────────────
_MAX_FAILURES      = 5
_LOCKOUT_SECONDS   = 15 * 60   # 15 minutes
_login_failures: dict[str, list] = defaultdict(list)   # ip → [datetime, ...]

def _is_locked(ip: str) -> bool:
    cutoff = datetime.now() - timedelta(seconds=_LOCKOUT_SECONDS)
    _login_failures[ip] = [t for t in _login_failures[ip] if t > cutoff]
    return len(_login_failures[ip]) >= _MAX_FAILURES

def _record_failure(ip: str):
    _login_failures[ip].append(datetime.now())

def _clear_failures(ip: str):
    _login_failures.pop(ip, None)


def _check_password(password: str) -> str | None:
    """Returns an error message if the password fails complexity rules, else None."""
    if len(password) < 8:
        return "Password must be at least 8 characters."
    if not re.search(r'[A-Z]', password):
        return "Password must contain at least one uppercase letter."
    if not re.search(r'[0-9]', password):
        return "Password must contain at least one digit."
    return None


def run_in_bg(fn, *args):
    """Run a blocking function in the default thread-pool so it doesn't stall the event loop."""
    loop = asyncio.get_running_loop()
    return loop.run_in_executor(None, fn, *args)

scheduler = BackgroundScheduler()
limiter   = Limiter(key_func=get_remote_address)


@asynccontextmanager
async def lifespan(app: FastAPI):
    # All startup jobs run in the background — server binds immediately without waiting for DB.
    scheduler.add_job(ensure_schema,              "date", id="startup_schema", next_run_time=datetime.now())
    scheduler.add_job(expire_validity_credits,    "date", id="startup_expire", next_run_time=datetime.now() + timedelta(seconds=15))

    scheduler.add_job(sync_sms_usage, "interval", hours=1, id="sms_sync",
                      next_run_time=datetime.now() + timedelta(seconds=30), max_instances=1)

    scheduler.add_job(notifier.send_daily_report, "cron", hour=18, minute=0, id="daily_report")
    scheduler.add_job(expire_validity_credits,    "cron", hour=0,  minute=0, id="validity_expire")
    scheduler.add_job(reset_daily_alert_flags,    "cron", hour=0,  minute=1, id="daily_flag_reset")

    scheduler.start()
    yield
    scheduler.shutdown(wait=False)   # don't block on in-flight jobs — DB handles its own transactions


app = FastAPI(lifespan=lifespan)
app.state.limiter = limiter
app.add_exception_handler(RateLimitExceeded, _rate_limit_exceeded_handler)

# ── Security headers — applied to every response ─────────────────────────────
class _SecurityHeaders(BaseHTTPMiddleware):
    async def dispatch(self, request: Request, call_next):
        response = await call_next(request)
        response.headers["X-Content-Type-Options"]  = "nosniff"
        response.headers["X-Frame-Options"]          = "DENY"
        response.headers["Referrer-Policy"]          = "strict-origin-when-cross-origin"
        response.headers["Content-Security-Policy"]  = (
            "default-src 'self'; "
            "script-src 'self' 'unsafe-inline' https://cdn.jsdelivr.net; "
            "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; "
            "font-src 'self' https://fonts.gstatic.com; "
            "img-src 'self' data:; "
            "connect-src 'self'"
        )
        if os.getenv("FORCE_HTTPS", "false").lower() == "true":
            response.headers["Strict-Transport-Security"] = "max-age=31536000; includeSubDomains"
        return response

app.add_middleware(_SecurityHeaders)

# ── HTTPS redirect — activate by setting FORCE_HTTPS=true in .env ────────────
if os.getenv("FORCE_HTTPS", "false").lower() == "true":
    from starlette.middleware.httpsredirect import HTTPSRedirectMiddleware
    app.add_middleware(HTTPSRedirectMiddleware)
    print("[STARTUP] HTTPS redirect ENABLED — all HTTP requests will be redirected to HTTPS.")

app.mount("/static", StaticFiles(directory=str(_BASE_DIR / "static")), name="static")
templates = Jinja2Templates(directory=str(_BASE_DIR / "templates"))

# ── Login page (GET) ──────────────────────────────────────
@app.get("/", response_class=HTMLResponse)
async def login_page(request: Request):
    return templates.TemplateResponse(
        request=request,
        name="Login.html",
        context={"error": None}
    )


# ── Login submit (POST) ───────────────────────────────────
@app.post("/login", response_class=HTMLResponse)
async def login(
    request: Request,
    username: str = Form(...),
    password: str = Form(...)
):
    ip = request.client.host if request.client else "unknown"

    if _is_locked(ip):
        log_audit(username, "LOGIN_BLOCKED", detail="Rate limited", ip_address=ip)
        failures     = _login_failures[ip]
        secs_left    = max(0, (min(failures) + timedelta(seconds=_LOCKOUT_SECONDS) - datetime.now()).total_seconds()) if failures else _LOCKOUT_SECONDS
        minutes_left = max(1, (int(secs_left) + 59) // 60)
        return templates.TemplateResponse(
            request=request,
            name="Login.html",
            context={"error": "Too many failed attempts. Access temporarily locked.", "lockout_minutes": minutes_left}
        )

    user = get_user_by_username(username)
    if not user or not verify_password(password, user["password_hash"]):
        _record_failure(ip)
        remaining = _MAX_FAILURES - len(_login_failures[ip])
        detail    = f"Invalid credentials ({remaining} attempt(s) left before lockout)"
        log_audit(username, "LOGIN_FAILED", detail=detail, ip_address=ip)
        return templates.TemplateResponse(
            request=request,
            name="Login.html",
            context={"error": "Invalid username or password.", "remaining": remaining}
        )

    if not user["is_active"]:
        log_audit(username, "LOGIN_FAILED", detail="Account inactive", ip_address=ip)
        return templates.TemplateResponse(
            request=request,
            name="Login.html",
            context={"error": "Your account is inactive. Contact the administrator."}
        )

    _clear_failures(ip)
    token = create_token(username, user["role"], user["token_version"])
    update_last_login(username)
    log_audit(username, "LOGIN", detail=f"role={user['role']}", ip_address=ip)

    response = RedirectResponse(url="/dashboard", status_code=303)
    response.set_cookie(
        key="access_token",
        value=token,
        httponly=True,
        samesite="strict",
        secure=os.getenv("FORCE_HTTPS", "false").lower() == "true",
        max_age=8 * 3600
    )
    return response


# ── Logout ────────────────────────────────────────────────
@app.get("/logout")
async def logout(request: Request):
    user = get_current_user(request)
    if user:
        ip = request.client.host if request.client else "unknown"
        bump_token_version(user["sub"])   # invalidate all existing tokens for this user
        log_audit(user["sub"], "LOGOUT", ip_address=ip)
    response = RedirectResponse(url="/", status_code=303)
    response.delete_cookie("access_token")
    return response


# ── Dashboard (GET) ───────────────────────────────────────
@app.get("/dashboard", response_class=HTMLResponse)
async def dashboard(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse(url="/", status_code=303)
    ip = request.client.host if request.client else "unknown"
    log_audit(user["sub"], "VIEW_DASHBOARD", ip_address=ip)
    return templates.TemplateResponse(
        request=request,
        name="Dashboard.html",
        context={"username": user["sub"], "role": user.get("role")}
    )


# ── User Management page (admin only) ────────────────────
@app.get("/users", response_class=HTMLResponse)
async def users_page(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse(url="/", status_code=303)
    if user.get("role") != "admin":
        return RedirectResponse(url="/dashboard", status_code=303)
    return templates.TemplateResponse(
        request=request,
        name="users.html",
        context={"username": user["sub"], "role": user.get("role")}
    )


# ── Audit Log page (admin only) ──────────────────────────
@app.get("/audit-log", response_class=HTMLResponse)
async def audit_log_page(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse(url="/", status_code=303)
    if user.get("role") != "admin":
        return RedirectResponse(url="/dashboard", status_code=303)
    return templates.TemplateResponse(
        request=request,
        name="audit_log.html",
        context={"username": user["sub"], "role": user.get("role")}
    )


# ── SMS Trends page (admin only) ─────────────────────────
@app.get("/sms-trends", response_class=HTMLResponse)
async def sms_trends_page(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse(url="/", status_code=303)
    if user.get("role") != "admin":
        return RedirectResponse(url="/dashboard", status_code=303)
    return templates.TemplateResponse(
        request=request,
        name="sms_trends.html",
        context={"username": user["sub"], "role": user.get("role")}
    )


# ── API: SMS trend data (admin only) ─────────────────────
@app.get("/api/sms-trend")
@limiter.limit("30/minute")
async def api_sms_trend(request: Request, period: str = "weekly",
                        _: dict = Depends(require_admin_api)):
    if period not in ("weekly", "monthly", "yearly"):
        return JSONResponse(
            content={"status": "error", "message": "period must be weekly, monthly, or yearly."},
            status_code=400
        )
    result = await run_in_bg(fetch_sms_trend, period)
    return JSONResponse(content={"status": "ok", **result})


# ── API: fetch real clients from DB ──────────────────────
@app.get("/api/clients")
@limiter.limit("60/minute")
async def get_clients(request: Request, _: dict = Depends(require_auth_api)):
    clients = await run_in_bg(fetch_clients)
    return JSONResponse(content={"status": "ok", "clients": clients})


# ── API: assign SMS credits (admin only) ─────────────────
@app.post("/api/assign-credits")
@limiter.limit("30/minute")
async def api_assign_credits(request: Request, data: dict = Body(...),
                              user: dict = Depends(require_admin_api)):
    try:
        client_id      = data.get("client_id")
        ip             = request.client.host if request.client else "unknown"
        start_date_str = (data.get("start_date") or "").strip()
        end_date_str   = (data.get("end_date")   or "").strip()

        try:
            allocated_sms = int(data.get("allocated_sms"))
        except (TypeError, ValueError):
            return JSONResponse(
                content={"status": "error", "message": "allocated_sms must be a valid integer."},
                status_code=400
            )

        if not client_id:
            return JSONResponse(
                content={"status": "error", "message": "client_id is required."},
                status_code=400
            )
        if allocated_sms <= 0:
            return JSONResponse(
                content={"status": "error", "message": "allocated_sms must be greater than zero."},
                status_code=400
            )
        if allocated_sms > 10_000_000:
            return JSONResponse(
                content={"status": "error", "message": "allocated_sms cannot exceed 10,000,000."},
                status_code=400
            )

        start_date = None
        end_date   = None
        if start_date_str or end_date_str:
            if not start_date_str or not end_date_str:
                return JSONResponse(
                    content={"status": "error", "message": "Both start_date and end_date are required when setting a validity period."},
                    status_code=400
                )
            try:
                from datetime import date as _date
                start_date = _date.fromisoformat(start_date_str)
                end_date   = _date.fromisoformat(end_date_str)
            except ValueError:
                return JSONResponse(
                    content={"status": "error", "message": "Invalid date format. Use YYYY-MM-DD."},
                    status_code=400
                )
            if end_date <= start_date:
                return JSONResponse(
                    content={"status": "error", "message": "End date must be after start date."},
                    status_code=400
                )

        client = await run_in_bg(assign_sms_credits, client_id, allocated_sms, start_date, end_date)

        if client:
            detail = f"allocated={allocated_sms}"
            if start_date:
                detail += f", validity={start_date} to {end_date}"
            log_audit(user["sub"], "ASSIGN_CREDITS", target_id=client_id,
                      detail=detail, ip_address=ip)
            return JSONResponse(content={"status": "ok", "client": client})

        return JSONResponse(
            content={"status": "error", "message": "Failed to save credits."},
            status_code=500
        )
    except Exception as e:
        import traceback
        traceback.print_exc()
        return JSONResponse(
            content={"status": "error", "message": "An internal error occurred. Check server logs."},
            status_code=500
        )


# ── API: client analytics ─────────────────────────────────
@app.get("/api/client-analytics/{client_id}")
@limiter.limit("60/minute")
async def api_client_analytics(request: Request, client_id: str,
                                user: dict = Depends(require_auth_api)):
    data = await run_in_bg(fetch_client_analytics, client_id)
    if not data:
        return JSONResponse(
            content={"status": "error", "message": "Client not found"},
            status_code=404
        )
    ip = request.client.host if request.client else "unknown"
    log_audit(user["sub"], "VIEW_CLIENT_ANALYTICS", target_id=client_id, ip_address=ip)
    return JSONResponse(content={"status": "ok", **data})


# ── API: dashboard data ───────────────────────────────────
@app.get("/api/dashboard-clients")
@limiter.limit("60/minute")
async def api_dashboard_clients(request: Request, force: bool = False, _: dict = Depends(require_auth_api)):
    if force:
        invalidate_dashboard_cache()
    clients = await run_in_bg(fetch_dashboard_clients)
    return JSONResponse(content={"status": "ok", "clients": clients})


# ── API: manually trigger SMS usage sync (admin only) ────
@app.post("/api/sync-usage")
@limiter.limit("5/minute")
async def api_sync_usage(request: Request, user: dict = Depends(require_admin_api)):
    result = await run_in_bg(sync_sms_usage)
    ip = request.client.host if request.client else "unknown"
    if result is None:
        return JSONResponse(
            content={"status": "error", "message": "Sync failed. Check server logs."},
            status_code=500
        )
    log_audit(user["sub"], "MANUAL_SYNC", detail=f"synced_rows={result['synced_rows']}", ip_address=ip)
    return JSONResponse(content={
        "status": "ok",
        "synced_rows":      result["synced_rows"],
        "affected_clients": result["affected_clients"]
    })


# ── API: threshold config ─────────────────────────────────
@app.get("/api/thresholds")
@limiter.limit("30/minute")
async def api_thresholds(request: Request, _: dict = Depends(require_auth_api)):
    return get_thresholds()


@app.post("/api/thresholds")
@limiter.limit("10/minute")
async def api_set_thresholds(request: Request, data: dict = Body(...),
                              user: dict = Depends(require_admin_api)):
    warning  = data.get("warning")
    critical = data.get("critical")
    ip       = request.client.host if request.client else "unknown"

    try:
        warning  = int(warning)
        critical = int(critical)
    except (TypeError, ValueError):
        return JSONResponse(
            content={"status": "error", "message": "warning and critical must be integers."},
            status_code=400
        )

    if not (1 <= warning <= 99) or not (1 <= critical <= 99):
        return JSONResponse(
            content={"status": "error", "message": "Values must be between 1 and 99."},
            status_code=400
        )
    if warning >= critical:
        return JSONResponse(
            content={"status": "error", "message": "Warning must be less than Critical."},
            status_code=400
        )

    ok = set_thresholds(warning, critical, updated_by=user["sub"])
    if not ok:
        return JSONResponse(
            content={"status": "error", "message": "Failed to save. Check server logs."},
            status_code=500
        )

    log_audit(user["sub"], "UPDATE_THRESHOLDS",
              detail=f"warning={warning} critical={critical}", ip_address=ip)
    return JSONResponse(content={"status": "ok", "warning": warning, "critical": critical})


# ── API: last sync timestamp ──────────────────────────────
@app.get("/api/last-synced")
@limiter.limit("60/minute")
async def api_last_synced(request: Request, _: dict = Depends(require_auth_api)):
    next_sync = None
    try:
        job = scheduler.get_job('sms_sync')
        if job and job.next_run_time:
            next_sync = job.next_run_time.isoformat()
    except Exception:
        pass
    return JSONResponse(content={"status": "ok", "last_synced": get_last_synced(), "next_sync": next_sync})


# ── API: fetch active alerts ──────────────────────────────
@app.get("/api/alerts")
@limiter.limit("60/minute")
async def api_alerts(request: Request, _: dict = Depends(require_auth_api)):
    alerts = await run_in_bg(fetch_active_alerts)
    return JSONResponse(content={"status": "ok", "alerts": alerts})


# ── API: send alert email (admin only) ───────────────────
@app.post("/api/send-alert-email")
@limiter.limit("20/minute")
async def api_send_alert_email(request: Request, data: dict = Body(...),
                                user: dict = Depends(require_admin_api)):
    client_id  = data.get("client_id")
    alert_type = data.get("alert_type")
    ip         = request.client.host if request.client else "unknown"

    if alert_type not in {"threshold_50", "threshold_90", "spike", "drop"}:
        return JSONResponse(
            content={"status": "error", "message": "Invalid alert_type."},
            status_code=400
        )
    if not client_id:
        return JSONResponse(
            content={"status": "error", "message": "client_id is required."},
            status_code=400
        )

    log_audit(user["sub"], "SEND_ALERT_EMAIL", target_id=client_id,
              detail=f"alert_type={alert_type}", ip_address=ip)
    # Fire-and-forget — SMTP is blocking I/O; respond instantly, send in background
    run_in_bg(notifier.send_for_client, client_id, alert_type)
    return JSONResponse(content={"status": "ok"})


# ── API: send consolidated report now (admin only) ───────
@app.post("/api/send-report-now")
@limiter.limit("5/minute")
async def api_send_report_now(request: Request, user: dict = Depends(require_admin_api)):
    ip = request.client.host if request.client else "unknown"
    log_audit(user["sub"], "SEND_REPORT_NOW", ip_address=ip)
    # Fire-and-forget — building Excel + SMTP takes several seconds; respond instantly
    run_in_bg(notifier.send_daily_report)
    return JSONResponse(content={"status": "ok"})


# ── API: send per-client status report (admin only) ──────
@app.post("/api/send-client-report")
@limiter.limit("20/minute")
async def api_send_client_report(request: Request, data: dict = Body(...),
                                  user: dict = Depends(require_admin_api)):
    client_id = data.get("client_id")
    ip        = request.client.host if request.client else "unknown"
    log_audit(user["sub"], "SEND_CLIENT_REPORT", target_id=client_id, ip_address=ip)
    ok = await run_in_bg(notifier.send_client_report, client_id)
    if not ok:
        return JSONResponse(content={"status": "error", "message": "Failed to send email. Check server logs."}, status_code=500)
    return JSONResponse(content={"status": "ok"})


# ── API: fetch saved client report emails ─────────────────
@app.get("/api/client-email/{client_id}")
@limiter.limit("60/minute")
async def api_get_client_email(client_id: str, request: Request,
                                _: dict = Depends(require_admin_api)):
    result = get_client_report_email(client_id)
    if result:
        return JSONResponse(content={"status": "ok", **result})
    return JSONResponse(content={"status": "not_found", "emails": []})


# ── API: save / update client report emails ───────────────
@app.post("/api/client-email")
@limiter.limit("30/minute")
async def api_save_client_email(request: Request, data: dict = Body(...),
                                 user: dict = Depends(require_admin_api)):
    client_id   = (data.get("client_id")   or "").strip()
    client_name = (data.get("client_name") or "").strip()
    emails_raw  = data.get("emails") or []
    ip          = request.client.host if request.client else "unknown"

    if not client_id:
        return JSONResponse(
            content={"status": "error", "message": "client_id is required."},
            status_code=400
        )
    if not isinstance(emails_raw, list):
        return JSONResponse(
            content={"status": "error", "message": "emails must be a list."},
            status_code=400
        )

    email_pattern = re.compile(r'^[^\s@]+@[^\s@]+\.[^\s@]+$')
    emails = [e.strip() for e in emails_raw if isinstance(e, str) and e.strip()]
    invalid = [e for e in emails if not email_pattern.match(e)]
    if invalid:
        return JSONResponse(
            content={"status": "error",
                     "message": f"Invalid email address(es): {', '.join(invalid)}"},
            status_code=400
        )

    ok = save_client_report_email(client_id, client_name, emails)
    if not ok:
        return JSONResponse(
            content={"status": "error", "message": "Failed to save emails. Check server logs."},
            status_code=500
        )
    log_audit(user["sub"], "SAVE_CLIENT_EMAIL", target_id=client_id,
              detail=f"emails={', '.join(emails)}", ip_address=ip)
    return JSONResponse(content={"status": "ok"})


# ── API: list all saved client emails (admin only) ────────
@app.get("/api/client-emails")
@limiter.limit("30/minute")
async def api_list_client_emails(request: Request, _: dict = Depends(require_admin_api)):
    emails = fetch_all_client_report_emails()
    return JSONResponse(content={"status": "ok", "emails": emails})


# ── Client Usage Report page (admin only) ─────────────────
@app.get("/client-usage-report", response_class=HTMLResponse)
async def client_usage_report_page(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse(url="/", status_code=303)
    if user.get("role") != "admin":
        return RedirectResponse(url="/dashboard", status_code=303)
    return templates.TemplateResponse(
        request=request,
        name="client_usage_report.html",
        context={"username": user["sub"], "role": user.get("role")}
    )


# ── API: day-wise SMS usage for a client (admin only) ─────
@app.get("/api/client-usage-report")
@limiter.limit("30/minute")
async def api_client_usage_report(
    request: Request,
    client_id:  str = "",
    start_date: str = "",
    end_date:   str = "",
    _: dict = Depends(require_admin_api)
):
    from datetime import date as _date
    client_id = client_id.strip()
    if not client_id or not start_date or not end_date:
        return JSONResponse(
            content={"status": "error", "message": "client_id, start_date, and end_date are required."},
            status_code=400
        )
    try:
        sd = _date.fromisoformat(start_date)
        ed = _date.fromisoformat(end_date)
    except ValueError:
        return JSONResponse(
            content={"status": "error", "message": "Invalid date format. Use YYYY-MM-DD."},
            status_code=400
        )
    if sd > ed:
        return JSONResponse(
            content={"status": "error", "message": "Start date cannot be after end date."},
            status_code=400
        )
    result = await run_in_bg(fetch_client_usage_report, client_id, sd, ed)
    if result is None:
        return JSONResponse(
            content={"status": "error", "message": "Client not found."},
            status_code=404
        )
    if not result["daily"]:
        return JSONResponse(content={"status": "empty", **result})
    return JSONResponse(content={"status": "ok", **result})


# ── API: download the DETAILED per-message report (SMS_Report_Template.xlsx
#         format — admin only) ─────────────────────────────────────────────
@app.get("/api/client-usage-report/download-detailed")
@limiter.limit("10/minute")
async def api_download_detailed_usage_report(
    request: Request,
    client_id: str = "",
    start_date: str = "",
    end_date: str = "",
    user: dict = Depends(require_admin_api),
):
    import io as _io

    from datetime import date as _date
    client_id = client_id.strip()
    if not client_id or not start_date or not end_date:
        return JSONResponse(
            content={"status": "error", "message": "Client, start date, and end date are required."},
            status_code=400,
        )
    try:
        sd = _date.fromisoformat(start_date)
        ed = _date.fromisoformat(end_date)
    except ValueError:
        return JSONResponse(
            content={"status": "error", "message": "Invalid date format. Use YYYY-MM-DD."},
            status_code=400,
        )
    if sd > ed:
        return JSONResponse(
            content={"status": "error", "message": "Start date cannot be after end date."},
            status_code=400,
        )

    output, filename, sms_count = await run_in_bg(
        generate_detailed_sms_report, client_id, start_date, end_date
    )
    if output is None:
        return JSONResponse(
            content={"status": "error", "message": "Unable to generate detailed SMS report."},
            status_code=500,
        )

    ip = request.client.host if request.client else "unknown"
    log_audit(
        user["sub"], "DOWNLOAD_DETAILED_SMS_REPORT", target_id=client_id,
        detail=f"from={start_date} to={end_date} sms_count={sms_count}", ip_address=ip,
    )
    return StreamingResponse(
        _io.BytesIO(output.getvalue()),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


# ── API: download usage report as Excel (admin only) ──────
@app.get("/api/client-usage-report/download")
@limiter.limit("10/minute")
async def api_download_usage_report(
    request: Request,
    client_id:  str = "",
    start_date: str = "",
    end_date:   str = "",
    user: dict = Depends(require_admin_api)
):
    import io as _io
    from datetime import date as _date
    client_id = client_id.strip()
    try:
        sd = _date.fromisoformat(start_date)
        ed = _date.fromisoformat(end_date)
    except ValueError:
        return JSONResponse(
            content={"status": "error", "message": "Invalid date format."},
            status_code=400
        )
    result = await run_in_bg(fetch_client_usage_report, client_id, sd, ed)
    if not result:
        return JSONResponse(
            content={"status": "error", "message": "Client not found."},
            status_code=404
        )
    buf      = await run_in_bg(build_usage_report_excel, result)
    safe_nm  = result["client_name"].replace(" ", "_")
    filename = f"sms_usage_{safe_nm}_{start_date}_{end_date}.xlsx"
    ip = request.client.host if request.client else "unknown"
    log_audit(user["sub"], "DOWNLOAD_USAGE_REPORT", target_id=client_id,
              detail=f"from={start_date} to={end_date}", ip_address=ip)
    return StreamingResponse(
        _io.BytesIO(buf.getvalue()),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )


# ── API: send usage report email to client (admin only) ───
@app.post("/api/client-usage-report/send-email")
@limiter.limit("10/minute")
async def api_send_usage_report_email(
    request: Request,
    data: dict = Body(...),
    user: dict = Depends(require_admin_api)
):
    from datetime import date as _date
    client_id  = (data.get("client_id")  or "").strip()
    start_date = (data.get("start_date") or "").strip()
    end_date   = (data.get("end_date")   or "").strip()
    ip         = request.client.host if request.client else "unknown"

    if not client_id or not start_date or not end_date:
        return JSONResponse(
            content={"status": "error", "message": "client_id, start_date, and end_date are required."},
            status_code=400
        )
    try:
        sd = _date.fromisoformat(start_date)
        ed = _date.fromisoformat(end_date)
    except ValueError:
        return JSONResponse(
            content={"status": "error", "message": "Invalid date format. Use YYYY-MM-DD."},
            status_code=400
        )
    if sd > ed:
        return JSONResponse(
            content={"status": "error", "message": "Start date cannot be after end date."},
            status_code=400
        )

    result = await run_in_bg(notifier.send_detailed_usage_report_email, client_id, sd, ed)

    if not result["ok"]:
        if result.get("reason") == "no_data":
            return JSONResponse(
                content={"status": "error", "message": "No usage data found for the selected date range."},
                status_code=400
            )
        return JSONResponse(
            content={"status": "error", "message": "Failed to send email. Check server logs."},
            status_code=500
        )

    log_audit(user["sub"], "SEND_USAGE_REPORT_EMAIL", target_id=client_id,
              detail=f"from={start_date} to={end_date} sms_count={result.get('sms_count', 0)}", ip_address=ip)
    return JSONResponse(content={
        "status": "ok",
        "emails": result.get("emails", []),
        "sms_count": result.get("sms_count", 0),
    })


# ── Client Email Management page (admin only) ─────────────
@app.get("/client-emails", response_class=HTMLResponse)
async def client_emails_page(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse(url="/", status_code=303)
    if user.get("role") != "admin":
        return RedirectResponse(url="/dashboard", status_code=303)
    return templates.TemplateResponse(
        request=request,
        name="client_emails.html",
        context={"username": user["sub"], "role": user.get("role")}
    )


# ── API: send report to saved client email ────────────────
@app.post("/api/send-report/{client_id}")
@limiter.limit("10/minute")
async def api_send_report_to_client(client_id: str, request: Request,
                                     user: dict = Depends(require_admin_api)):
    ip     = request.client.host if request.client else "unknown"
    result = await run_in_bg(notifier.send_report_to_client_email, client_id)

    if not result["ok"]:
        if result.get("reason") == "no_data":
            return JSONResponse(
                content={"status": "error", "message": "Client not found in database."},
                status_code=404
            )
        return JSONResponse(
            content={"status": "error", "message": "Failed to send report. Check server logs."},
            status_code=500
        )

    log_audit(user["sub"], "SEND_CLIENT_REPORT_EMAIL", target_id=client_id, ip_address=ip)
    return JSONResponse(content={"status": "ok"})


# ── API: list all users (admin only) ──────────────────────
@app.get("/api/users")
@limiter.limit("30/minute")
async def api_list_users(request: Request, _: dict = Depends(require_admin_api)):
    users = fetch_users()
    return JSONResponse(content={"status": "ok", "users": users})


# ── API: create user (admin only) ─────────────────────────
@app.post("/api/users")
@limiter.limit("10/minute")
async def api_create_user(request: Request, data: dict = Body(...),
                           user: dict = Depends(require_admin_api)):
    username = (data.get("username") or "").strip()
    password = (data.get("password") or "").strip()
    role     = (data.get("role") or "viewer").strip()
    ip       = request.client.host if request.client else "unknown"

    if not username or not password:
        return JSONResponse(
            content={"status": "error", "message": "Username and password are required."},
            status_code=400
        )
    if role not in ("admin", "viewer"):
        return JSONResponse(
            content={"status": "error", "message": "Role must be 'admin' or 'viewer'."},
            status_code=400
        )
    pw_err = _check_password(password)
    if pw_err:
        return JSONResponse(content={"status": "error", "message": pw_err}, status_code=400)

    result = create_user(username, password, role)
    if result is None:
        return JSONResponse(
            content={"status": "error", "message": "Database error — user was not created."},
            status_code=500
        )
    if "error" in result:
        return JSONResponse(
            content={"status": "error", "message": result["error"]},
            status_code=409
        )

    log_audit(user["sub"], "CREATE_USER", target_id=username,
              detail=f"role={role}", ip_address=ip)
    return JSONResponse(content={"status": "ok", "user": result}, status_code=201)


# ── API: activate / deactivate user (admin only) ──────────
@app.patch("/api/users/{username}/status")
@limiter.limit("10/minute")
async def api_set_user_status(username: str, request: Request,
                               data: dict = Body(...),
                               user: dict = Depends(require_admin_api)):
    is_active = data.get("is_active")
    if is_active is None:
        return JSONResponse(
            content={"status": "error", "message": "'is_active' field required."},
            status_code=400
        )
    ip = request.client.host if request.client else "unknown"

    if not is_active and username == user["sub"]:
        return JSONResponse(
            content={"status": "error", "message": "You cannot deactivate your own account."},
            status_code=400
        )

    result = set_user_status(username, bool(is_active))
    if "error" in result:
        return JSONResponse(
            content={"status": "error", "message": result["error"]},
            status_code=400
        )

    action = "ACTIVATE_USER" if is_active else "DEACTIVATE_USER"
    log_audit(user["sub"], action, target_id=username, ip_address=ip)
    return JSONResponse(content={"status": "ok"})


# ── API: reset user password (admin only) ─────────────────
@app.post("/api/users/{username}/reset-password")
@limiter.limit("10/minute")
async def api_reset_password(username: str, request: Request,
                              data: dict = Body(...),
                              user: dict = Depends(require_admin_api)):
    new_password = (data.get("new_password") or "").strip()
    ip           = request.client.host if request.client else "unknown"

    pw_err = _check_password(new_password)
    if pw_err:
        return JSONResponse(content={"status": "error", "message": pw_err}, status_code=400)

    ok = reset_user_password(username, new_password)
    if not ok:
        return JSONResponse(
            content={"status": "error", "message": "Failed to reset password. Check server logs."},
            status_code=500
        )

    log_audit(user["sub"], "RESET_PASSWORD", target_id=username, ip_address=ip)
    return JSONResponse(content={"status": "ok"})


# ── API: audit log (admin only) ───────────────────────────
@app.get("/api/audit-log")
@limiter.limit("30/minute")
async def api_audit_log(
    request: Request,
    limit:  int = 50,
    offset: int = 0,
    user:   str = "",
    action: str = "",
    _: dict = Depends(require_admin_api)
):
    result = fetch_audit_log(
        limit        = min(max(limit, 1), 200),
        offset       = max(offset, 0),
        user_filter  = user.strip()   or None,
        action_filter= action.strip() or None,
    )
    return JSONResponse(content={"status": "ok", **result})