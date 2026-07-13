"""
CAMERA_APP.PY -- BRHAS Dashcam Camera Control Web App (v2)
==========================================================
Flask app for toggling Motive dashcam cameras on personal vehicles.

Features:
  - Multi-user PIN auth (admin / dispatch / driver roles)
  - Role-based vehicle visibility
  - PostgreSQL for persistent state and activity logging
  - Double-click protection
  - Master log + per-user log views

Deploy to Render as a separate service from brhas-safety-api.

Env vars required:
  DATABASE_URL   - PostgreSQL connection string (from Render)
  MOTIVE_API_KEY - Motive API key
  SECRET_KEY     - Flask session secret
"""

import os
import time
import secrets
from datetime import datetime, timezone
from functools import wraps
from urllib.parse import quote as url_quote

import psycopg2
import psycopg2.extras
import requests
from flask import Flask, request, jsonify, session, redirect, url_for

# =============================================================================
# CONFIG
# =============================================================================
DATABASE_URL = os.environ.get("DATABASE_URL", "")
MOTIVE_API_KEY = os.environ.get("MOTIVE_API_KEY", "")
MOTIVE_BASE = "https://api.gomotive.com/v1"

MOTIVE_HEADERS = {
    "X-Api-Key": MOTIVE_API_KEY,
    "Accept": "application/json",
    "Content-Type": "application/json",
}

# =============================================================================
# APP SETUP
# =============================================================================
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", secrets.token_hex(32))
app.permanent_session_lifetime = 86400  # 24 hours


# =============================================================================
# DATABASE
# =============================================================================
def get_db():
    """Get a database connection."""
    conn = psycopg2.connect(DATABASE_URL)
    conn.autocommit = True
    return conn


def db_query(sql, params=None, fetchone=False, fetchall=False):
    """Execute a query and optionally fetch results."""
    conn = get_db()
    try:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(sql, params)
            if fetchone:
                return cur.fetchone()
            if fetchall:
                return cur.fetchall()
            return None
    finally:
        conn.close()


def db_execute(sql, params=None):
    """Execute a write query."""
    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute(sql, params)
    finally:
        conn.close()


# =============================================================================
# AUTH HELPERS
# =============================================================================
def get_current_user():
    """Get the current logged-in user from session."""
    user_id = session.get("user_id")
    if not user_id:
        return None
    return db_query("SELECT * FROM users WHERE id = %s", (user_id,), fetchone=True)


def require_auth(f):
    """Require any authenticated user."""
    @wraps(f)
    def decorated(*args, **kwargs):
        user = get_current_user()
        if not user:
            return redirect(url_for("login_page"))
        return f(*args, user=user, **kwargs)
    return decorated


def require_role(*roles):
    """Require specific role(s)."""
    def decorator(f):
        @wraps(f)
        def decorated(*args, **kwargs):
            user = get_current_user()
            if not user:
                return redirect(url_for("login_page"))
            if user["role"] not in roles:
                return "Access denied", 403
            return f(*args, user=user, **kwargs)
        return decorated
    return decorator


def get_visible_vehicles(user):
    """Get vehicles this user is allowed to see."""
    if user["role"] == "admin":
        return db_query("SELECT * FROM vehicles ORDER BY location, vehicle_number",
                        fetchall=True)
    elif user["role"] == "dispatch":
        return db_query(
            "SELECT * FROM vehicles WHERE location = %s ORDER BY vehicle_number",
            (user["location"],), fetchall=True)
    else:  # driver
        return db_query(
            "SELECT * FROM vehicles WHERE driver_name = %s ORDER BY vehicle_number",
            (user["name"],), fetchall=True)


def can_toggle_vehicle(user, vehicle_number):
    """Check if user is allowed to toggle this vehicle."""
    if user["role"] == "admin":
        return True
    vehicle = db_query("SELECT * FROM vehicles WHERE vehicle_number = %s",
                       (vehicle_number,), fetchone=True)
    if not vehicle:
        return False
    if user["role"] == "dispatch":
        return vehicle["location"] == user["location"]
    # driver
    return vehicle["driver_name"] == user["name"]


def get_visible_log(user, limit=100):
    """Get activity log entries this user is allowed to see."""
    if user["role"] == "admin":
        return db_query(
            "SELECT * FROM activity_log ORDER BY timestamp DESC LIMIT %s",
            (limit,), fetchall=True)
    elif user["role"] == "dispatch":
        vehicles = get_visible_vehicles(user)
        vnums = [v["vehicle_number"] for v in vehicles]
        if not vnums:
            return []
        return db_query(
            "SELECT * FROM activity_log WHERE vehicle_number = ANY(%s) "
            "ORDER BY timestamp DESC LIMIT %s",
            (vnums, limit), fetchall=True)
    else:  # driver
        return db_query(
            "SELECT * FROM activity_log WHERE user_name = %s "
            "ORDER BY timestamp DESC LIMIT %s",
            (user["name"], limit), fetchall=True)


# =============================================================================
# MOTIVE API
# =============================================================================
def motive_toggle(eld_device_id, state):
    """Send camera ON/OFF command to Motive. Returns (req_id, req_status, error)."""
    url = f"{MOTIVE_BASE}/cameras/{eld_device_id}"
    try:
        resp = requests.put(url, headers=MOTIVE_HEADERS,
                            json={"camera_state": state.upper()}, timeout=15)
        if resp.status_code == 429:
            return None, None, "Rate limited by Motive. Try again in 30 seconds."
        resp.raise_for_status()
        data = resp.json()
        return data.get("req_id"), data.get("req_status"), None
    except requests.HTTPError as e:
        msg = e.response.text if e.response else str(e)
        return None, None, f"Motive API error: {msg}"
    except Exception as e:
        return None, None, f"Connection error: {str(e)}"


def motive_poll(eld_device_id, req_id, max_attempts=4):
    """Poll for toggle completion. Returns final status string."""
    url = f"{MOTIVE_BASE}/cameras/{eld_device_id}/{req_id}"
    for attempt in range(max_attempts):
        time.sleep(5)
        try:
            resp = requests.get(url, headers=MOTIVE_HEADERS, timeout=10)
            if resp.status_code == 429:
                time.sleep(15)
                continue
            resp.raise_for_status()
            status = resp.json().get("req_status", "Unknown")
            if status in ("Succeeded", "Failed", "Error"):
                return status
        except Exception:
            continue
    return "Pending"


# =============================================================================
# DEDUP (prevent double-click)
# =============================================================================
_recent_toggles = {}  # vehicle_number -> timestamp


def is_duplicate_toggle(vehicle_number, cooldown=5):
    """Return True if this vehicle was toggled within cooldown seconds."""
    now = time.time()
    last = _recent_toggles.get(vehicle_number, 0)
    if now - last < cooldown:
        return True
    _recent_toggles[vehicle_number] = now
    return False


# =============================================================================
# SHARED CSS + JS
# =============================================================================
SHARED_CSS = """
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background:#0f172a;color:#e2e8f0;min-height:100vh}
.header{background:#1e293b;border-bottom:2px solid #334155;padding:16px 24px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px}
.header h1{font-size:20px}
.header .sub{color:#94a3b8;font-size:13px}
.header-right{display:flex;align-items:center;gap:16px}
.header-right .user-info{color:#94a3b8;font-size:13px}
.header-right .role-badge{background:#334155;color:#94a3b8;padding:2px 8px;border-radius:10px;font-size:11px;text-transform:uppercase}
.logout{color:#94a3b8;text-decoration:none;font-size:13px}
.logout:hover{color:#e2e8f0}
.container{max-width:1200px;margin:0 auto;padding:20px}
.stats{display:flex;gap:16px;margin-bottom:20px;flex-wrap:wrap}
.stat{background:#1e293b;border:1px solid #334155;border-radius:8px;padding:12px 16px}
.stat .label{font-size:11px;color:#64748b;text-transform:uppercase}
.stat .val{font-size:24px;font-weight:700}
.tbl-wrap{background:#1e293b;border-radius:8px;border:1px solid #334155;overflow-x:auto;margin-bottom:24px}
table{width:100%;border-collapse:collapse}
th{background:#334155;padding:10px 14px;text-align:left;font-size:12px;text-transform:uppercase;color:#94a3b8;font-weight:600}
td{padding:10px 14px;border-top:1px solid #0f172a;font-size:14px}
tr:hover td{background:#253349}
h2{font-size:16px;margin-bottom:12px;color:#94a3b8}
.bulk{margin-bottom:16px;display:flex;gap:8px;flex-wrap:wrap}
.bulk form{display:inline}
.bulk button{padding:8px 16px;border-radius:6px;border:none;font-size:14px;font-weight:600;cursor:pointer}
.nav-tabs{display:flex;gap:4px;margin-bottom:20px}
.nav-tabs a{color:#94a3b8;text-decoration:none;padding:8px 16px;border-radius:6px;font-size:14px}
.nav-tabs a:hover{background:#1e293b;color:#e2e8f0}
.nav-tabs a.active{background:#334155;color:#e2e8f0;font-weight:600}
"""

DOUBLE_CLICK_JS = """
<script>
document.querySelectorAll('form.toggle-form').forEach(function(form) {
  form.addEventListener('submit', function(e) {
    var btn = form.querySelector('button');
    if (btn.disabled) { e.preventDefault(); return; }
    btn.disabled = true;
    btn.textContent = 'Sending...';
    btn.style.opacity = '0.5';
  });
});
</script>
"""


# =============================================================================
# HTML TEMPLATES
# =============================================================================
def login_html(error=None):
    err = f'<div style="color:#fca5a5;margin-bottom:12px;font-size:14px">{error}</div>' if error else ""
    return f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Login - BRHAS Camera Control</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background:#0f172a;color:#e2e8f0;min-height:100vh;display:flex;align-items:center;justify-content:center}}
.card{{background:#1e293b;border-radius:12px;border:1px solid #334155;padding:32px;width:100%;max-width:360px;text-align:center}}
h1{{font-size:20px;margin-bottom:4px}}
.sub{{color:#94a3b8;font-size:14px;margin-bottom:24px}}
input{{width:100%;background:#0f172a;border:1px solid #475569;border-radius:8px;padding:12px;color:#e2e8f0;font-size:18px;text-align:center;letter-spacing:8px;margin-bottom:16px}}
input:focus{{outline:none;border-color:#3b82f6}}
button{{width:100%;background:#3b82f6;color:white;border:none;border-radius:8px;padding:12px;font-size:16px;font-weight:600;cursor:pointer}}
button:hover{{background:#2563eb}}
</style></head>
<body><div class="card">
<h1>Camera Control</h1>
<div class="sub">BRHAS - Enter your PIN</div>
{err}
<form method="POST" action="/login">
<input type="password" name="pin" placeholder="PIN" maxlength="10" autofocus inputmode="numeric">
<button type="submit">Enter</button>
</form>
</div></body></html>"""


def header_html(user, active_tab="dashboard"):
    role_colors = {"admin": "#3b82f6", "dispatch": "#22c55e", "driver": "#f59e0b"}
    rc = role_colors.get(user["role"], "#94a3b8")
    loc = f" - {user['location']}" if user.get("location") else ""

    tabs = f'<a href="/dashboard" class="{"active" if active_tab == "dashboard" else ""}">Dashboard</a>'
    tabs += f'<a href="/log" class="{"active" if active_tab == "log" else ""}">Activity Log</a>'

    return f"""<div class="header">
<div><h1>Camera Control</h1><div class="sub">BRHAS - Personal Vehicles</div></div>
<div class="header-right">
<span class="user-info">{user['name']}{loc}</span>
<span class="role-badge" style="color:{rc};border:1px solid {rc}">{user['role']}</span>
<a href="/logout" class="logout">Logout</a>
</div></div>
<div class="container"><div class="nav-tabs">{tabs}</div>"""


def dashboard_html(user, vehicles, log_entries):
    on_count = sum(1 for v in vehicles if v["camera_state"] == "ON")
    off_count = sum(1 for v in vehicles if v["camera_state"] == "OFF")

    rows = ""
    for v in vehicles:
        state = v["camera_state"] or "ON"
        sc = "#86efac" if state == "ON" else "#fca5a5"
        badge_bg = "#166534" if state == "ON" else "#7f1d1d"
        toggled = ""
        if v.get("last_toggled_at"):
            toggled = v["last_toggled_at"].strftime("%m/%d %I:%M %p") if hasattr(v["last_toggled_at"], "strftime") else str(v["last_toggled_at"])
            if v.get("last_toggled_by"):
                toggled += f" by {v['last_toggled_by']}"
        rows += f"""<tr>
<td><strong>{v['vehicle_number']}</strong></td>
<td>{v['driver_name']}</td>
<td>{v['location']}</td>
<td>{v['department']}</td>
<td><span style="background:{badge_bg};color:{sc};padding:2px 10px;border-radius:12px;font-size:12px;font-weight:600">{state}</span></td>
<td style="font-size:12px;color:#64748b">{toggled}</td>
<td style="white-space:nowrap">
<form method="POST" action="/api/toggle" class="toggle-form" style="display:inline">
<input type="hidden" name="vehicle_number" value="{v['vehicle_number']}">
<input type="hidden" name="state" value="ON">
<button type="submit" style="background:#166534;color:#86efac;border:1px solid #22c55e;border-radius:6px;padding:4px 12px;cursor:pointer;font-size:12px;font-weight:600">ON</button>
</form>
<form method="POST" action="/api/toggle" class="toggle-form" style="display:inline">
<input type="hidden" name="vehicle_number" value="{v['vehicle_number']}">
<input type="hidden" name="state" value="OFF">
<button type="submit" style="background:#7f1d1d;color:#fca5a5;border:1px solid #ef4444;border-radius:6px;padding:4px 12px;cursor:pointer;font-size:12px;font-weight:600">OFF</button>
</form>
</td></tr>"""

    log_rows = ""
    for entry in log_entries[:20]:
        ts = entry["timestamp"].strftime("%m/%d %I:%M %p") if hasattr(entry["timestamp"], "strftime") else str(entry["timestamp"])
        ac = "#86efac" if entry["action"] == "ON" else "#fca5a5"
        log_rows += f"""<tr>
<td style="color:#64748b">{ts}</td>
<td>{entry['vehicle_number']}</td>
<td>{entry['user_name']}</td>
<td><span style="color:{ac}">{entry['action']}</span></td>
<td style="color:#94a3b8">{entry.get('motive_status', '')}</td></tr>"""

    bulk_html = ""
    if user["role"] in ("admin", "dispatch"):
        bulk_html = """<div class="bulk">
<form method="POST" action="/api/bulk-toggle" class="toggle-form">
<input type="hidden" name="state" value="ON">
<button type="submit" style="background:#166534;color:#86efac" onclick="return confirm('Turn ALL visible cameras ON?')">All Cameras ON</button>
</form>
<form method="POST" action="/api/bulk-toggle" class="toggle-form">
<input type="hidden" name="state" value="OFF">
<button type="submit" style="background:#7f1d1d;color:#fca5a5" onclick="return confirm('Turn ALL visible cameras OFF?')">All Cameras OFF</button>
</form></div>"""

    return f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Dashboard - BRHAS Camera Control</title>
<style>{SHARED_CSS}</style></head>
<body>
{header_html(user, "dashboard")}
<div class="stats">
<div class="stat"><div class="label">Total Vehicles</div><div class="val">{len(vehicles)}</div></div>
<div class="stat"><div class="label">Cameras ON</div><div class="val" style="color:#86efac">{on_count}</div></div>
<div class="stat"><div class="label">Cameras OFF</div><div class="val" style="color:#fca5a5">{off_count}</div></div>
</div>
{bulk_html}
<div class="tbl-wrap"><table>
<thead><tr><th>Vehicle #</th><th>Driver</th><th>Location</th><th>Dept</th><th>Camera</th><th>Last Changed</th><th>Actions</th></tr></thead>
<tbody>{rows}</tbody>
</table></div>

<h2>Recent Activity</h2>
<div class="tbl-wrap"><table>
<thead><tr><th>Time</th><th>Vehicle</th><th>User</th><th>Action</th><th>Status</th></tr></thead>
<tbody>{log_rows}</tbody>
</table></div>
<div style="text-align:center;margin-top:8px"><a href="/log" style="color:#3b82f6;text-decoration:none;font-size:13px">View full activity log</a></div>
</div>
{DOUBLE_CLICK_JS}
</body></html>"""


def log_page_html(user, log_entries, filter_user=None):
    # User filter dropdown (admin only)
    user_filter = ""
    if user["role"] == "admin":
        users = db_query("SELECT DISTINCT user_name FROM activity_log ORDER BY user_name", fetchall=True)
        options = '<option value="">All Users</option>'
        for u in (users or []):
            selected = 'selected' if filter_user == u["user_name"] else ''
            options += f'<option value="{u["user_name"]}" {selected}>{u["user_name"]}</option>'
        user_filter = f"""<div style="margin-bottom:16px">
<form method="GET" action="/log" style="display:flex;gap:8px;align-items:center">
<label style="color:#94a3b8;font-size:13px">Filter by user:</label>
<select name="user" onchange="this.form.submit()" style="background:#1e293b;border:1px solid #475569;border-radius:6px;padding:6px 10px;color:#e2e8f0;font-size:14px">{options}</select>
</form></div>"""

    log_rows = ""
    for entry in log_entries:
        ts = entry["timestamp"].strftime("%Y-%m-%d %I:%M:%S %p") if hasattr(entry["timestamp"], "strftime") else str(entry["timestamp"])
        ac = "#86efac" if entry["action"] == "ON" else "#fca5a5"
        log_rows += f"""<tr>
<td style="color:#64748b;white-space:nowrap">{ts}</td>
<td>{entry['vehicle_number']}</td>
<td>{entry['user_name']}</td>
<td><span style="background:#334155;padding:2px 8px;border-radius:10px;font-size:11px">{entry.get('user_role', '')}</span></td>
<td><span style="color:{ac};font-weight:600">{entry['action']}</span></td>
<td style="color:#64748b;font-family:monospace;font-size:12px">{entry.get('motive_req_id', '') or ''}</td>
<td style="color:#94a3b8">{entry.get('motive_status', '') or ''}</td></tr>"""

    return f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Activity Log - BRHAS Camera Control</title>
<style>{SHARED_CSS}</style></head>
<body>
{header_html(user, "log")}
{user_filter}
<div class="tbl-wrap"><table>
<thead><tr><th>Time</th><th>Vehicle</th><th>User</th><th>Role</th><th>Action</th><th>Request ID</th><th>Status</th></tr></thead>
<tbody>{log_rows}</tbody>
</table></div>
</div></body></html>"""


# =============================================================================
# ROUTES
# =============================================================================
@app.route("/")
def index():
    if get_current_user():
        return redirect(url_for("dashboard"))
    return redirect(url_for("login_page"))


@app.route("/login", methods=["GET"])
def login_page():
    if get_current_user():
        return redirect(url_for("dashboard"))
    return login_html()


@app.route("/login", methods=["POST"])
def login_post():
    pin = request.form.get("pin", "").strip()
    if not pin:
        return login_html("Please enter your PIN"), 401

    user = db_query("SELECT * FROM users WHERE pin = %s", (pin,), fetchone=True)
    if not user:
        return login_html("Invalid PIN"), 401

    session["user_id"] = user["id"]
    session.permanent = True
    return redirect(url_for("dashboard"))


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login_page"))


@app.route("/dashboard")
@require_auth
def dashboard(user):
    vehicles = get_visible_vehicles(user)
    log_entries = get_visible_log(user, limit=20)
    return dashboard_html(user, vehicles or [], log_entries or [])


@app.route("/log")
@require_auth
def log_view(user):
    filter_user = request.args.get("user", "").strip() or None

    if user["role"] == "admin" and filter_user:
        log_entries = db_query(
            "SELECT * FROM activity_log WHERE user_name = %s ORDER BY timestamp DESC LIMIT 500",
            (filter_user,), fetchall=True)
    else:
        log_entries = get_visible_log(user, limit=500)

    return log_page_html(user, log_entries or [], filter_user)


# -- API --
@app.route("/api/toggle", methods=["POST"])
@require_auth
def api_toggle(user):
    if request.is_json:
        data = request.get_json()
    else:
        data = request.form

    vehicle_number = data.get("vehicle_number", "").strip()
    state = data.get("state", "").upper().strip()

    if state not in ("ON", "OFF"):
        if request.is_json:
            return jsonify({"error": "State must be ON or OFF"}), 400
        return redirect(url_for("dashboard"))

    # Permission check
    if not can_toggle_vehicle(user, vehicle_number):
        if request.is_json:
            return jsonify({"error": "Access denied"}), 403
        return redirect(url_for("dashboard"))

    # Double-click protection
    if is_duplicate_toggle(vehicle_number):
        if request.is_json:
            return jsonify({"error": "Request already processing"}), 429
        return redirect(url_for("dashboard"))

    # Get vehicle
    vehicle = db_query("SELECT * FROM vehicles WHERE vehicle_number = %s",
                       (vehicle_number,), fetchone=True)
    if not vehicle:
        if request.is_json:
            return jsonify({"error": "Vehicle not found"}), 404
        return redirect(url_for("dashboard"))

    eld_id = vehicle["eld_device_id"]

    # Call Motive API
    req_id, req_status, error = motive_toggle(eld_id, state)

    if error:
        if request.is_json:
            return jsonify({"error": error}), 502
        return redirect(url_for("dashboard"))

    # Poll for confirmation
    motive_status = req_status or "Unknown"
    if req_id and req_status == "Submitted":
        final = motive_poll(eld_id, req_id, max_attempts=3)
        motive_status = final

    # Update vehicle state in DB
    db_execute(
        "UPDATE vehicles SET camera_state = %s, last_toggled_at = %s, last_toggled_by = %s "
        "WHERE vehicle_number = %s",
        (state, datetime.now(timezone.utc), user["name"], vehicle_number))

    # Log the action
    db_execute(
        "INSERT INTO activity_log (timestamp, user_name, user_role, vehicle_number, action, motive_req_id, motive_status) "
        "VALUES (%s, %s, %s, %s, %s, %s, %s)",
        (datetime.now(timezone.utc), user["name"], user["role"], vehicle_number,
         state, req_id, motive_status))

    if request.is_json:
        return jsonify({
            "vehicle_number": vehicle_number,
            "state": state,
            "req_id": req_id,
            "motive_status": motive_status,
        })

    return redirect(url_for("dashboard"))


@app.route("/api/bulk-toggle", methods=["POST"])
@require_role("admin", "dispatch")
def api_bulk_toggle(user):
    state = request.form.get("state", "").upper()
    if state not in ("ON", "OFF"):
        return redirect(url_for("dashboard"))

    vehicles = get_visible_vehicles(user)
    for v in (vehicles or []):
        eld_id = v["eld_device_id"]
        req_id, req_status, error = motive_toggle(eld_id, state)
        motive_status = "Error" if error else (req_status or "Unknown")

        if not error:
            db_execute(
                "UPDATE vehicles SET camera_state = %s, last_toggled_at = %s, last_toggled_by = %s "
                "WHERE vehicle_number = %s",
                (state, datetime.now(timezone.utc), user["name"], v["vehicle_number"]))

        db_execute(
            "INSERT INTO activity_log (timestamp, user_name, user_role, vehicle_number, action, motive_req_id, motive_status) "
            "VALUES (%s, %s, %s, %s, %s, %s, %s)",
            (datetime.now(timezone.utc), user["name"], user["role"], v["vehicle_number"],
             state, req_id, motive_status))

        time.sleep(2)  # Rate limit protection

    return redirect(url_for("dashboard"))


@app.route("/api/health")
def health():
    db_ok = False
    try:
        db_query("SELECT 1", fetchone=True)
        db_ok = True
    except Exception:
        pass
    return jsonify({
        "status": "ok" if db_ok else "db_error",
        "database": db_ok,
        "motive_key_set": bool(MOTIVE_API_KEY),
    })


# =============================================================================
if __name__ == "__main__":
    app.run(debug=True, port=5001)
