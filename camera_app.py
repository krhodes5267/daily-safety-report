"""
CAMERA_APP.PY -- BRHAS Dashcam Camera Control Web App
=====================================================
Flask app for toggling Motive dashcam cameras on personal vehicles.

Two views:
  - Driver page:   /v/<vehicle_number>  (no PIN, one-button toggle)
  - Dispatch page: /dispatch            (PIN-protected, all vehicles)

Deploy to Render as a separate service from brhas-safety-api.

Env vars required:
  MOTIVE_API_KEY  - Motive API key
  DISPATCH_PIN    - PIN for dispatch dashboard access
  SECRET_KEY      - Flask session secret (auto-generated if not set)
"""

import os
import time
import secrets
from datetime import datetime, timezone
from functools import wraps

import requests
from flask import Flask, request, jsonify, session, redirect, url_for

# =============================================================================
# CONFIG
# =============================================================================
MOTIVE_API_KEY = os.environ.get("MOTIVE_API_KEY", "")
DISPATCH_PIN = os.environ.get("DISPATCH_PIN", "1234")
MOTIVE_BASE = "https://api.gomotive.com/v1"

MOTIVE_HEADERS = {
    "X-Api-Key": MOTIVE_API_KEY,
    "Accept": "application/json",
    "Content-Type": "application/json",
}

# Personal vehicles list -- replace with real vehicles once they're in Motive
# For now, using 2 company vehicles for testing
PERSONAL_VEHICLES = [
    {"eld_device_id": "1674558", "vehicle_number": "2294C", "driver_name": "Bob Stokes", "location": "Midland", "department": "Field"},
    {"eld_device_id": "1680021", "vehicle_number": "2135C", "driver_name": "Jose Gonzalez", "location": "Laredo", "department": "Field"},
    # -------------------------------------------------------------------------
    # ADD REAL PERSONAL VEHICLES BELOW (once added to Motive with dashcams)
    # Format: {"eld_device_id": "XXXXXX", "vehicle_number": "XXXXX",
    #          "driver_name": "First Last", "location": "Yard", "department": "Dept"}
    # -------------------------------------------------------------------------
]

# Build lookup dicts
VEHICLES_BY_NUMBER = {v["vehicle_number"]: v for v in PERSONAL_VEHICLES}
VEHICLES_BY_ELD = {v["eld_device_id"]: v for v in PERSONAL_VEHICLES}

# =============================================================================
# APP SETUP
# =============================================================================
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", secrets.token_hex(32))

# In-memory state (resets on deploy — acceptable for toggle state since
# the real state lives in Motive, this is just for display)
camera_state = {}   # vehicle_number -> "ON" | "OFF"
activity_log = []   # [{timestamp, vehicle, driver, action, source}]

MAX_LOG_ENTRIES = 500


def log_action(vehicle_number, action, source):
    """Record a toggle action."""
    driver = VEHICLES_BY_NUMBER.get(vehicle_number, {}).get("driver_name", "Unknown")
    entry = {
        "timestamp": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"),
        "vehicle": vehicle_number,
        "driver": driver,
        "action": action,
        "source": source,
    }
    activity_log.insert(0, entry)
    if len(activity_log) > MAX_LOG_ENTRIES:
        activity_log[:] = activity_log[:MAX_LOG_ENTRIES]


def require_dispatch_pin(f):
    """Decorator to require dispatch PIN for a route."""
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("dispatch_auth"):
            return redirect(url_for("dispatch_login"))
        return f(*args, **kwargs)
    return decorated


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


def motive_poll(eld_device_id, req_id, max_attempts=6):
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
# HTML TEMPLATES (inline to keep it single-file)
# =============================================================================
def driver_page_html(vehicle, state, message=None, msg_type="info"):
    bg_on = "#166534" if state == "ON" else "#1e293b"
    bg_off = "#7f1d1d" if state == "OFF" else "#1e293b"
    border_on = "#22c55e" if state == "ON" else "#475569"
    border_off = "#ef4444" if state == "OFF" else "#475569"
    status_color = "#86efac" if state == "ON" else "#fca5a5"
    status_text = "Recording" if state == "ON" else "Cameras Off"
    msg_html = ""
    if message:
        mc = "#86efac" if msg_type == "success" else "#fca5a5" if msg_type == "error" else "#fde047"
        msg_html = f'<div style="background:#1e293b;border:1px solid {mc};border-radius:8px;padding:12px;margin-bottom:20px;color:{mc};font-size:14px">{message}</div>'

    return f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0,user-scalable=no">
<title>{vehicle['vehicle_number']} - Camera Control</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background:#0f172a;color:#e2e8f0;min-height:100vh;display:flex;flex-direction:column;align-items:center;padding:20px}}
.card{{background:#1e293b;border-radius:16px;border:1px solid #334155;padding:24px;width:100%;max-width:400px;text-align:center}}
.vehicle-num{{font-size:28px;font-weight:700;margin-bottom:4px}}
.driver{{color:#94a3b8;font-size:16px;margin-bottom:4px}}
.location{{color:#64748b;font-size:14px;margin-bottom:20px}}
.status{{font-size:18px;font-weight:600;color:{status_color};margin-bottom:24px}}
.status-dot{{display:inline-block;width:12px;height:12px;border-radius:50%;background:{status_color};margin-right:8px;animation:{('pulse 2s infinite' if state=='ON' else 'none')}}}
@keyframes pulse{{0%,100%{{opacity:1}}50%{{opacity:.4}}}}
.btn-row{{display:flex;gap:12px;justify-content:center}}
.btn{{flex:1;padding:16px;border-radius:12px;border:2px solid;font-size:18px;font-weight:600;cursor:pointer;transition:all .15s}}
.btn-on{{background:{bg_on};color:#86efac;border-color:{border_on}}}
.btn-off{{background:{bg_off};color:#fca5a5;border-color:{border_off}}}
.btn:active{{transform:scale(.96)}}
.btn:disabled{{opacity:.5;cursor:not-allowed}}
form{{display:inline;flex:1}}
.branding{{color:#475569;font-size:12px;margin-top:20px}}
</style></head>
<body>
<div class="card">
  <div class="vehicle-num">{vehicle['vehicle_number']}</div>
  <div class="driver">{vehicle['driver_name']}</div>
  <div class="location">{vehicle['location']} - {vehicle['department']}</div>
  <div class="status"><span class="status-dot"></span>{status_text}</div>
  {msg_html}
  <div class="btn-row">
    <form method="POST" action="/api/toggle">
      <input type="hidden" name="vehicle_number" value="{vehicle['vehicle_number']}">
      <input type="hidden" name="state" value="ON">
      <input type="hidden" name="source" value="driver">
      <button type="submit" class="btn btn-on" style="width:100%">ON</button>
    </form>
    <form method="POST" action="/api/toggle">
      <input type="hidden" name="vehicle_number" value="{vehicle['vehicle_number']}">
      <input type="hidden" name="state" value="OFF">
      <input type="hidden" name="source" value="driver">
      <button type="submit" class="btn btn-off" style="width:100%">OFF</button>
    </form>
  </div>
</div>
<div class="branding">BRHAS Camera Control</div>
</body></html>"""


def dispatch_login_html(error=None):
    err = f'<div style="color:#fca5a5;margin-bottom:12px;font-size:14px">{error}</div>' if error else ""
    return f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Dispatch Login - Camera Control</title>
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
<h1>Dispatch Dashboard</h1>
<div class="sub">BRHAS Camera Control</div>
{err}
<form method="POST" action="/dispatch/login">
<input type="password" name="pin" placeholder="PIN" maxlength="10" autofocus>
<button type="submit">Enter</button>
</form>
</div></body></html>"""


def dispatch_dashboard_html(vehicles_data, log_data):
    rows = ""
    for v in vehicles_data:
        state = v.get("state", "Unknown")
        sc = "#86efac" if state == "ON" else "#fca5a5" if state == "OFF" else "#94a3b8"
        badge_bg = "#166534" if state == "ON" else "#7f1d1d" if state == "OFF" else "#334155"
        rows += f"""<tr>
<td><strong>{v['vehicle_number']}</strong></td>
<td>{v['driver_name']}</td>
<td>{v['location']}</td>
<td>{v['department']}</td>
<td><span style="background:{badge_bg};color:{sc};padding:2px 10px;border-radius:12px;font-size:12px;font-weight:600">{state}</span></td>
<td style="white-space:nowrap">
<form method="POST" action="/api/toggle" style="display:inline">
<input type="hidden" name="vehicle_number" value="{v['vehicle_number']}">
<input type="hidden" name="state" value="ON"><input type="hidden" name="source" value="dispatch">
<button type="submit" style="background:#166534;color:#86efac;border:1px solid #22c55e;border-radius:6px;padding:4px 12px;cursor:pointer;font-size:12px;font-weight:600">ON</button>
</form>
<form method="POST" action="/api/toggle" style="display:inline">
<input type="hidden" name="vehicle_number" value="{v['vehicle_number']}">
<input type="hidden" name="state" value="OFF"><input type="hidden" name="source" value="dispatch">
<button type="submit" style="background:#7f1d1d;color:#fca5a5;border:1px solid #ef4444;border-radius:6px;padding:4px 12px;cursor:pointer;font-size:12px;font-weight:600">OFF</button>
</form>
</td></tr>"""

    log_rows = ""
    for entry in log_data[:50]:
        log_rows += f"""<tr>
<td style="color:#64748b">{entry['timestamp']}</td>
<td>{entry['vehicle']}</td>
<td>{entry['driver']}</td>
<td><span style="color:{'#86efac' if entry['action']=='ON' else '#fca5a5'}">{entry['action']}</span></td>
<td style="color:#94a3b8">{entry['source']}</td></tr>"""

    on_count = sum(1 for v in vehicles_data if v.get("state") == "ON")
    off_count = sum(1 for v in vehicles_data if v.get("state") == "OFF")

    return f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Dispatch - Camera Control</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background:#0f172a;color:#e2e8f0;min-height:100vh}}
.header{{background:#1e293b;border-bottom:2px solid #334155;padding:16px 24px;display:flex;align-items:center;justify-content:space-between}}
.header h1{{font-size:20px}}
.header .sub{{color:#94a3b8;font-size:13px}}
.logout{{color:#94a3b8;text-decoration:none;font-size:13px}}
.logout:hover{{color:#e2e8f0}}
.container{{max-width:1200px;margin:0 auto;padding:20px}}
.stats{{display:flex;gap:16px;margin-bottom:20px;flex-wrap:wrap}}
.stat{{background:#1e293b;border:1px solid #334155;border-radius:8px;padding:12px 16px}}
.stat .label{{font-size:11px;color:#64748b;text-transform:uppercase}}
.stat .val{{font-size:24px;font-weight:700}}
.bulk{{margin-bottom:16px;display:flex;gap:8px}}
.bulk form{{display:inline}}
.bulk button{{padding:8px 16px;border-radius:6px;border:none;font-size:14px;font-weight:600;cursor:pointer}}
.tbl-wrap{{background:#1e293b;border-radius:8px;border:1px solid #334155;overflow-x:auto;margin-bottom:24px}}
table{{width:100%;border-collapse:collapse}}
th{{background:#334155;padding:10px 14px;text-align:left;font-size:12px;text-transform:uppercase;color:#94a3b8;font-weight:600}}
td{{padding:10px 14px;border-top:1px solid #0f172a;font-size:14px}}
tr:hover td{{background:#253349}}
h2{{font-size:16px;margin-bottom:12px;color:#94a3b8}}
</style></head>
<body>
<div class="header">
<div><h1>Dispatch Dashboard</h1><div class="sub">BRHAS Camera Control - Personal Vehicles</div></div>
<a href="/dispatch/logout" class="logout">Logout</a>
</div>
<div class="container">
<div class="stats">
<div class="stat"><div class="label">Total Vehicles</div><div class="val">{len(vehicles_data)}</div></div>
<div class="stat"><div class="label">Cameras ON</div><div class="val" style="color:#86efac">{on_count}</div></div>
<div class="stat"><div class="label">Cameras OFF</div><div class="val" style="color:#fca5a5">{off_count}</div></div>
</div>

<div class="bulk">
<form method="POST" action="/api/bulk-toggle">
<input type="hidden" name="state" value="ON">
<button type="submit" style="background:#166534;color:#86efac" onclick="return confirm('Turn ALL cameras ON?')">All Cameras ON</button>
</form>
<form method="POST" action="/api/bulk-toggle">
<input type="hidden" name="state" value="OFF">
<button type="submit" style="background:#7f1d1d;color:#fca5a5" onclick="return confirm('Turn ALL cameras OFF?')">All Cameras OFF</button>
</form>
</div>

<div class="tbl-wrap"><table>
<thead><tr><th>Vehicle #</th><th>Driver</th><th>Location</th><th>Department</th><th>Camera</th><th>Actions</th></tr></thead>
<tbody>{rows}</tbody>
</table></div>

<h2>Activity Log</h2>
<div class="tbl-wrap"><table>
<thead><tr><th>Time</th><th>Vehicle</th><th>Driver</th><th>Action</th><th>Source</th></tr></thead>
<tbody>{log_rows}</tbody>
</table></div>
</div></body></html>"""


# =============================================================================
# ROUTES
# =============================================================================
@app.route("/")
def index():
    return redirect(url_for("dispatch_login"))


# -- Driver View --
@app.route("/v/<vehicle_number>")
def driver_view(vehicle_number):
    vehicle = VEHICLES_BY_NUMBER.get(vehicle_number)
    if not vehicle:
        return f"<h1>Vehicle {vehicle_number} not found</h1>", 404
    state = camera_state.get(vehicle_number, "ON")
    msg = request.args.get("msg")
    msg_type = request.args.get("msg_type", "info")
    return driver_page_html(vehicle, state, msg, msg_type)


# -- Dispatch --
@app.route("/dispatch/login", methods=["GET"])
def dispatch_login():
    if session.get("dispatch_auth"):
        return redirect(url_for("dispatch_dashboard"))
    return dispatch_login_html()


@app.route("/dispatch/login", methods=["POST"])
def dispatch_login_post():
    pin = request.form.get("pin", "")
    if pin == DISPATCH_PIN:
        session["dispatch_auth"] = True
        session.permanent = True
        return redirect(url_for("dispatch_dashboard"))
    return dispatch_login_html("Incorrect PIN"), 401


@app.route("/dispatch/logout")
def dispatch_logout():
    session.pop("dispatch_auth", None)
    return redirect(url_for("dispatch_login"))


@app.route("/dispatch")
@require_dispatch_pin
def dispatch_dashboard():
    vehicles_data = []
    for v in PERSONAL_VEHICLES:
        vehicles_data.append({
            **v,
            "state": camera_state.get(v["vehicle_number"], "ON"),
        })
    return dispatch_dashboard_html(vehicles_data, activity_log)


# -- API --
@app.route("/api/toggle", methods=["POST"])
def api_toggle():
    # Accept both form data and JSON
    if request.is_json:
        data = request.get_json()
    else:
        data = request.form

    vehicle_number = data.get("vehicle_number", "").strip()
    state = data.get("state", "").upper().strip()
    source = data.get("source", "api")

    if vehicle_number not in VEHICLES_BY_NUMBER:
        if request.is_json:
            return jsonify({"error": f"Vehicle {vehicle_number} not found"}), 404
        return f"Vehicle {vehicle_number} not found", 404

    if state not in ("ON", "OFF"):
        if request.is_json:
            return jsonify({"error": "State must be ON or OFF"}), 400
        return "State must be ON or OFF", 400

    vehicle = VEHICLES_BY_NUMBER[vehicle_number]
    eld_id = vehicle["eld_device_id"]

    # Call Motive API
    req_id, req_status, error = motive_toggle(eld_id, state)

    if error:
        if request.is_json:
            return jsonify({"error": error}), 502
        # Redirect back with error message
        if source == "driver":
            return redirect(f"/v/{vehicle_number}?msg={error}&msg_type=error")
        return redirect(f"/dispatch?error={error}")

    # Update local state
    camera_state[vehicle_number] = state
    log_action(vehicle_number, state, source)

    # Poll for confirmation (non-blocking for web, just try once)
    if req_id and req_status == "Submitted":
        final = motive_poll(eld_id, req_id, max_attempts=3)
        confirmed = final == "Succeeded"
    else:
        confirmed = False

    if request.is_json:
        return jsonify({
            "vehicle_number": vehicle_number,
            "state": state,
            "req_id": req_id,
            "confirmed": confirmed,
        })

    # Redirect back to source page
    if confirmed:
        msg = f"Camera {state} confirmed"
        msg_type = "success"
    else:
        msg = f"Camera {state} command sent (may take a moment to apply)"
        msg_type = "info"

    if source == "driver":
        return redirect(f"/v/{vehicle_number}?msg={msg}&msg_type={msg_type}")
    return redirect(url_for("dispatch_dashboard"))


@app.route("/api/bulk-toggle", methods=["POST"])
@require_dispatch_pin
def api_bulk_toggle():
    state = request.form.get("state", "").upper()
    if state not in ("ON", "OFF"):
        return redirect(url_for("dispatch_dashboard"))

    for v in PERSONAL_VEHICLES:
        eld_id = v["eld_device_id"]
        req_id, req_status, error = motive_toggle(eld_id, state)
        if not error:
            camera_state[v["vehicle_number"]] = state
            log_action(v["vehicle_number"], state, "dispatch-bulk")
        # Small delay between API calls to avoid rate limiting
        time.sleep(2)

    return redirect(url_for("dispatch_dashboard"))


@app.route("/api/vehicles")
def api_vehicles():
    result = []
    for v in PERSONAL_VEHICLES:
        result.append({
            **v,
            "state": camera_state.get(v["vehicle_number"], "ON"),
        })
    return jsonify(result)


@app.route("/api/log")
def api_log():
    return jsonify(activity_log[:100])


@app.route("/api/health")
def health():
    return jsonify({
        "status": "ok",
        "vehicles_configured": len(PERSONAL_VEHICLES),
        "motive_key_set": bool(MOTIVE_API_KEY),
    })


# =============================================================================
if __name__ == "__main__":
    app.run(debug=True, port=5001)
