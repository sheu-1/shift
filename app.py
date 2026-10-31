import io
import os
import re
import json
import random
import ssl
import urllib.request
import urllib.error
from datetime import datetime, date, timedelta
from types import SimpleNamespace
import sqlite3

from flask import (
    Flask, render_template, request, redirect, url_for, flash, send_file, jsonify
)
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "change-this-secret-key-in-production")

# Try loading from .env if present
env_path = os.path.join(app.root_path, ".env")
if not os.path.exists(env_path):
    # Check .env.example
    example_path = os.path.join(app.root_path, ".env.example")
    if os.path.exists(example_path):
        env_path = example_path

if os.path.exists(env_path):
    try:
        with open(env_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    k, v = line.split("=", 1)
                    k, v = k.strip(), v.strip().strip("'\"")
                    if k not in os.environ:
                        os.environ[k] = v
    except Exception:
        pass

SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_SERVICE_KEY")

supabase = None
if SUPABASE_URL and SUPABASE_KEY:
    try:
        from supabase import create_client
        supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
    except Exception as e:
        print(f"Supabase init error (falling back to SQLite): {e}")
        supabase = None

# SQLite Database setup
INSTANCE_DIR = os.path.join(app.root_path, "instance")
DB_PATH = os.path.join(INSTANCE_DIR, "shifts.db")

def get_db_conn():
    os.makedirs(INSTANCE_DIR, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

_DEFAULT_SHIFTS = [
    dict(code="AM",    label="AM Shift",    start_time="6:00 AM",  end_time="3:00 PM",  capacity=3, sort_order=0),
    dict(code="SWING", label="Swing Shift", start_time="11:00 AM", end_time="8:00 PM",  capacity=1, sort_order=1),
    dict(code="PM",    label="PM Shift",    start_time="3:00 PM",  end_time="12:00 AM", capacity=3, sort_order=2),
]

def init_db():
    conn = get_db_conn()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS worker (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name VARCHAR(120) NOT NULL,
            shift VARCHAR(10),
            added_at DATETIME,
            leave_days VARCHAR(100) DEFAULT ''
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS shift_config (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code VARCHAR(10) UNIQUE NOT NULL,
            label VARCHAR(50) NOT NULL,
            start_time VARCHAR(20) NOT NULL,
            end_time VARCHAR(20) NOT NULL,
            capacity INTEGER NOT NULL DEFAULT 1,
            sort_order INTEGER NOT NULL DEFAULT 0
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS assignment (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            worker_id INTEGER NOT NULL,
            day VARCHAR(15) NOT NULL,
            shift_code VARCHAR(10)
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS app_config (
            key TEXT PRIMARY KEY,
            value TEXT
        )
    """)
    count = conn.execute("SELECT count(*) FROM shift_config").fetchone()[0]
    if count == 0:
        for s in _DEFAULT_SHIFTS:
            conn.execute(
                "INSERT INTO shift_config (code, label, start_time, end_time, capacity, sort_order) VALUES (?, ?, ?, ?, ?, ?)",
                (s["code"], s["label"], s["start_time"], s["end_time"], s["capacity"], s["sort_order"])
            )
    conn.commit()
    conn.close()

    if supabase:
        try:
            res = supabase.table("shift_config").select("id", count="exact").execute()
            if res.count == 0:
                supabase.table("shift_config").insert(_DEFAULT_SHIFTS).execute()
        except Exception as e:
            print(f"Error seeding Supabase shifts: {e}")

init_db()

# ------------------------------------------------------------- app settings helper --
def get_setting(key, default=None):
    env_val = os.environ.get(key.upper())
    if env_val:
        return env_val
    conn = get_db_conn()
    r = conn.execute("SELECT value FROM app_config WHERE key = ?", (key,)).fetchone()
    conn.close()
    if r and r["value"] is not None:
        return r["value"]
    return default

def set_setting(key, value):
    conn = get_db_conn()
    conn.execute("INSERT OR REPLACE INTO app_config (key, value) VALUES (?, ?)", (key, str(value)))
    conn.commit()
    conn.close()

# ------------------------------------------------------------- config & dates --
def get_roster_date_range():
    if supabase:
        try:
            res = supabase.table("worker").select("name").ilike("name", "___CONFIG___|%").execute()
            if res.data:
                name_val = res.data[0].get("name", "")
                parts = name_val.split("|")
                if len(parts) == 3:
                    return parts[1], parts[2]
        except Exception as e:
            print(f"Error fetching config from Supabase: {e}")
    # SQLite fallback
    conn = get_db_conn()
    r = conn.execute("SELECT value FROM app_config WHERE key = 'roster_date_range'").fetchone()
    if not r:
        w = conn.execute("SELECT name FROM worker WHERE name LIKE '___CONFIG___|%'").fetchone()
        if w:
            parts = w["name"].split("|")
            if len(parts) == 3:
                conn.close()
                return parts[1], parts[2]
    conn.close()
    if r and r["value"]:
        parts = r["value"].split("|")
        if len(parts) == 2:
            return parts[0], parts[1]
    return None, None

def set_roster_date_range(start_date, end_date):
    if supabase:
        try:
            res = supabase.table("worker").select("id").ilike("name", "___CONFIG___|%").execute()
            if res.data:
                supabase.table("worker").update({"name": f"___CONFIG___|{start_date}|{end_date}"}).eq("id", res.data[0]["id"]).execute()
            else:
                supabase.table("worker").insert({"name": f"___CONFIG___|{start_date}|{end_date}"}).execute()
        except Exception as e:
            print(f"Error saving config in Supabase: {e}")
    # SQLite
    conn = get_db_conn()
    conn.execute("INSERT OR REPLACE INTO app_config (key, value) VALUES ('roster_date_range', ?)", (f"{start_date}|{end_date}",))
    w = conn.execute("SELECT id FROM worker WHERE name LIKE '___CONFIG___|%'").fetchone()
    if w:
        conn.execute("UPDATE worker SET name = ? WHERE id = ?", (f"___CONFIG___|{start_date}|{end_date}", w["id"]))
    else:
        conn.execute("INSERT INTO worker (name, leave_days) VALUES (?, '')", (f"___CONFIG___|{start_date}|{end_date}",))
    conn.commit()
    conn.close()

def get_dynamic_days():
    start_str, end_str = get_roster_date_range()
    today = date.today()
    if start_str and end_str:
        try:
            start_d = datetime.strptime(start_str, "%Y-%m-%d").date()
            end_d = datetime.strptime(end_str, "%Y-%m-%d").date()
        except ValueError:
            start_d, end_d = today, today + timedelta(days=6)
    else:
        start_d, end_d = today, today + timedelta(days=6)
    
    if end_d < start_d:
        end_d = start_d
        
    days = []
    curr = start_d
    while curr <= end_d:
        days.append(curr.strftime("%Y-%m-%d"))
        curr += timedelta(days=1)
    return days

def format_day_short(d_str):
    try:
        d = datetime.strptime(d_str, "%Y-%m-%d")
        return d.strftime("%d %a")
    except:
        return d_str

# ------------------------------------------------------------- DB Data Helpers --
def get_shift_configs():
    if supabase:
        try:
            res = supabase.table("shift_config").select("*").order("sort_order").execute()
            return [SimpleNamespace(**d) for d in res.data]
        except Exception as e:
            print(f"Error fetching shift configs from Supabase: {e}")
    conn = get_db_conn()
    rows = conn.execute("SELECT * FROM shift_config ORDER BY sort_order").fetchall()
    conn.close()
    return [SimpleNamespace(**dict(r)) for r in rows]

def shift_meta_dict():
    return {s.code: s for s in get_shift_configs()}

def worker_leave_set(worker):
    """Return set of leave day names for a worker."""
    if not getattr(worker, 'leave_days', None):
        return set()
    return {d.strip() for d in worker.leave_days.split(",") if d.strip()}

def get_all_workers_with_assignments():
    if supabase:
        try:
            res = supabase.table("worker").select("*, assignments:assignment(*)").not_.ilike("name", "___CONFIG___|%").order("name").execute()
            workers = []
            for d in res.data:
                w = SimpleNamespace(**d)
                w.assignments = [SimpleNamespace(**a) for a in getattr(w, 'assignments', [])]
                workers.append(w)
            return workers
        except Exception as e:
            print(f"Error fetching workers from Supabase: {e}")
    conn = get_db_conn()
    workers_rows = conn.execute("SELECT * FROM worker WHERE name NOT LIKE '___CONFIG___|%' ORDER BY name").fetchall()
    assignments_rows = conn.execute("SELECT * FROM assignment").fetchall()
    conn.close()
    assign_map = {}
    for a in assignments_rows:
        wid = a["worker_id"]
        if wid not in assign_map:
            assign_map[wid] = []
        assign_map[wid].append(SimpleNamespace(**dict(a)))
    workers = []
    for w in workers_rows:
        obj = SimpleNamespace(**dict(w))
        obj.assignments = assign_map.get(obj.id, [])
        workers.append(obj)
    return workers

def name_exists(name):
    clean = name.strip()
    if supabase:
        try:
            res = supabase.table("worker").select("id").ilike("name", clean).execute()
            return len(res.data) > 0
        except Exception as e:
            print(f"Supabase name_exists error: {e}")
    conn = get_db_conn()
    r = conn.execute("SELECT id FROM worker WHERE LOWER(name) = LOWER(?)", (clean,)).fetchone()
    conn.close()
    return r is not None

def add_worker_db(name):
    clean = name.strip()
    if supabase:
        try:
            supabase.table("worker").insert({"name": clean, "leave_days": ""}).execute()
            return True
        except Exception as e:
            print(f"Supabase add_worker error: {e}")
    conn = get_db_conn()
    conn.execute("INSERT INTO worker (name, leave_days) VALUES (?, '')", (clean,))
    conn.commit()
    conn.close()
    return True

def delete_worker_db(worker_id):
    if supabase:
        try:
            res = supabase.table("worker").select("name").eq("id", worker_id).execute()
            name = res.data[0]["name"] if res.data else "Worker"
            supabase.table("assignment").delete().eq("worker_id", worker_id).execute()
            supabase.table("worker").delete().eq("id", worker_id).execute()
            return name
        except Exception as e:
            print(f"Supabase delete_worker error: {e}")
    conn = get_db_conn()
    r = conn.execute("SELECT name FROM worker WHERE id = ?", (worker_id,)).fetchone()
    name = r["name"] if r else "Worker"
    conn.execute("DELETE FROM assignment WHERE worker_id = ?", (worker_id,))
    conn.execute("DELETE FROM worker WHERE id = ?", (worker_id,))
    conn.commit()
    conn.close()
    return name

def update_worker_leave_db(worker_id, leave_str):
    if supabase:
        try:
            supabase.table("worker").update({"leave_days": leave_str}).eq("id", worker_id).execute()
            return
        except Exception as e:
            print(f"Supabase update leave error: {e}")
    conn = get_db_conn()
    conn.execute("UPDATE worker SET leave_days = ? WHERE id = ?", (leave_str, worker_id))
    conn.commit()
    conn.close()

def update_shift_config_db(cfg_id, label, start_time, end_time, capacity):
    if supabase:
        try:
            supabase.table("shift_config").update({
                "label": label,
                "start_time": start_time,
                "end_time": end_time,
                "capacity": capacity
            }).eq("id", cfg_id).execute()
            return
        except Exception as e:
            print(f"Supabase update shift_config error: {e}")
    conn = get_db_conn()
    conn.execute(
        "UPDATE shift_config SET label = ?, start_time = ?, end_time = ?, capacity = ? WHERE id = ?",
        (label, start_time, end_time, capacity, cfg_id)
    )
    conn.commit()
    conn.close()

def assign_single_shift_db(worker_id, day, shift_code):
    if shift_code == "" or shift_code == "—":
        shift_code = None

    if supabase:
        try:
            res = supabase.table("assignment").select("id").eq("worker_id", worker_id).eq("day", day).execute()
            if res.data:
                row_id = res.data[0]["id"]
                supabase.table("assignment").update({"shift_code": shift_code}).eq("id", row_id).execute()
            else:
                supabase.table("assignment").insert({"worker_id": worker_id, "day": day, "shift_code": shift_code}).execute()
            return True
        except Exception as e:
            print(f"Supabase assign_single_shift error: {e}")

    conn = get_db_conn()
    r = conn.execute("SELECT id FROM assignment WHERE worker_id = ? AND day = ?", (worker_id, day)).fetchone()
    if r:
        conn.execute("UPDATE assignment SET shift_code = ? WHERE id = ?", (shift_code, r["id"]))
    else:
        conn.execute("INSERT INTO assignment (worker_id, day, shift_code) VALUES (?, ?, ?)", (worker_id, day, shift_code))
    conn.commit()
    conn.close()
    return True

def delete_assignments_for_days_db(days_list):
    if not days_list:
        return
    if supabase:
        try:
            supabase.table("assignment").delete().in_("day", days_list).execute()
            return
        except Exception as e:
            print(f"Supabase delete_assignments error: {e}")
    conn = get_db_conn()
    placeholders = ",".join(["?"] * len(days_list))
    conn.execute(f"DELETE FROM assignment WHERE day IN ({placeholders})", tuple(days_list))
    conn.commit()
    conn.close()

def insert_assignments_db(new_assignments):
    if not new_assignments:
        return
    if supabase:
        try:
            supabase.table("assignment").insert(new_assignments).execute()
            return
        except Exception as e:
            print(f"Supabase insert_assignments error: {e}")
    conn = get_db_conn()
    for a in new_assignments:
        conn.execute(
            "INSERT INTO assignment (worker_id, day, shift_code) VALUES (?, ?, ?)",
            (a["worker_id"], a["day"], a.get("shift_code"))
        )
    conn.commit()
    conn.close()

def grouped_workers_for_day(day_str):
    """Return (groups, unassigned, off_workers, leave_workers, sl_workers, ph_workers, all_workers) for a given date."""
    configs    = get_shift_configs()
    codes      = [c.code for c in configs]
    all_workers = get_all_workers_with_assignments()

    groups        = {code: [] for code in codes}
    unassigned    = []
    off_workers   = []
    leave_workers = []
    sl_workers    = []
    ph_workers    = []

    for w in all_workers:
        assign = next((a for a in w.assignments if a.day == day_str), None)
        if assign:
            if assign.shift_code in groups:
                groups[assign.shift_code].append(w)
            elif assign.shift_code == "OFF":
                off_workers.append(w)
            elif assign.shift_code == "LEAVE":
                leave_workers.append(w)
            elif assign.shift_code in ("SL", "SO"):
                sl_workers.append(w)
            elif assign.shift_code == "PH":
                ph_workers.append(w)
            else:
                unassigned.append(w)
        else:
            try:
                dt = datetime.strptime(day_str, "%Y-%m-%d")
                weekday_name = dt.strftime("%A")
            except:
                weekday_name = day_str
            if weekday_name in worker_leave_set(w):
                leave_workers.append(w)
            else:
                unassigned.append(w)

    return groups, unassigned, off_workers, leave_workers, sl_workers, ph_workers, all_workers

# ------------------------------------------------------------- Core Constraints & Auto-allocate --
def is_baby_care_worker(name):
    """Rule 2: Jacky and Eunice (Equnice) have young babies at home -> AM and SWING only."""
    n = (name or "").lower()
    return "jacky" in n or "jackie" in n or "eunice" in n or "equnice" in n

def run_auto_allocation():
    """
    Allocates shifts enforcing all 4 strict system prompt rules:
    1. Off-to-PM: When a staff is coming from OFF, they ONLY enter PM shifts.
    2. Young babies (Jacky & Eunice): ONLY enter AM and SWING shifts, never PM.
    3. Balanced PM and AM shifts across active staff over the roster.
    4. Complete coverage: Every staff member is allocated a shift given any day.
       If not on shift, they are on OFF, LEAVE, SL, or PH (no blank/None cells).
    """
    workers = get_all_workers_with_assignments()
    if not workers:
        return []

    configs = get_shift_configs()
    days_list = get_dynamic_days()
    if not days_list:
        return []

    # Shift capacities
    config_map = {c.code: c for c in configs}
    am_cap = config_map.get("AM", SimpleNamespace(capacity=3)).capacity
    swing_cap = config_map.get("SWING", SimpleNamespace(capacity=1)).capacity
    pm_cap = config_map.get("PM", SimpleNamespace(capacity=3)).capacity

    # Wipe existing assignments for the current date range
    delete_assignments_for_days_db(days_list)

    workers_list = list(workers)
    leave_map = {w.id: worker_leave_set(w) for w in workers_list}

    schedule = {}  # (worker_id, day) -> shift_code
    tallies = {w.id: {"AM": 0, "PM": 0, "SWING": 0, "OFF": 0, "LEAVE": 0} for w in workers_list}

    # Step 1: Pre-assign LEAVE rows
    for w in workers_list:
        for d in days_list:
            try:
                weekday_name = datetime.strptime(d, "%Y-%m-%d").strftime("%A")
            except:
                weekday_name = d
            if weekday_name in leave_map[w.id]:
                schedule[(w.id, d)] = "LEAVE"
                tallies[w.id]["LEAVE"] += 1

    # Step 2: Fairly distribute 1-2 OFF days per worker across the week
    num_days = len(days_list)
    target_offs = max(1, num_days // 7)
    daily_off_counts = {d: 0 for d in days_list}

    shuffled_workers = list(workers_list)
    random.shuffle(shuffled_workers)

    for w in shuffled_workers:
        avail_for_off = [d for d in days_list if (w.id, d) not in schedule]
        if not avail_for_off:
            continue
        # Pick day with lowest current off count; tiebreak with random
        avail_for_off.sort(key=lambda d: (daily_off_counts[d], random.random()))
        for off_day in avail_for_off[:target_offs]:
            schedule[(w.id, off_day)] = "OFF"
            tallies[w.id]["OFF"] += 1
            daily_off_counts[off_day] += 1

    # Step 3: Day-by-day chronological shift assignment
    for idx, d in enumerate(days_list):
        prev_day = days_list[idx - 1] if idx > 0 else None

        # Workers available today
        unassigned_today = [w for w in workers_list if (w.id, d) not in schedule]
        if not unassigned_today:
            continue

        # Rule 1: Coming from OFF -> Only PM
        coming_from_off = set()
        if prev_day:
            for w in unassigned_today:
                if schedule.get((w.id, prev_day)) == "OFF":
                    coming_from_off.add(w.id)

        # Rule 2: Baby care workers (Jacky & Eunice) -> Only AM & SWING
        baby_care_ids = {w.id for w in unassigned_today if is_baby_care_worker(w.name)}

        # Allocate PM Shifts first:
        # Eligible for PM = not baby care workers
        pm_eligible = [w for w in unassigned_today if w.id not in baby_care_ids]

        # Prioritize workers coming from OFF (mandatory PM entrance)
        must_pm = [w for w in pm_eligible if w.id in coming_from_off]
        other_pm = [w for w in pm_eligible if w.id not in coming_from_off]

        # Rule 3: Balance PM shifts
        must_pm.sort(key=lambda w: (tallies[w.id]["PM"], random.random()))
        other_pm.sort(key=lambda w: (tallies[w.id]["PM"], random.random()))
        pm_pool = must_pm + other_pm

        pm_chosen = pm_pool[:pm_cap]
        for w in pm_chosen:
            schedule[(w.id, d)] = "PM"
            tallies[w.id]["PM"] += 1
            unassigned_today = [u for u in unassigned_today if u.id != w.id]

        # Allocate AM Shifts:
        # Eligible for AM = not coming from OFF (unless baby care worker whose restriction is AM/SWING)
        am_eligible = [w for w in unassigned_today if (w.id not in coming_from_off or w.id in baby_care_ids)]
        # Rule 3: Balance AM shifts & prioritize baby care workers in AM
        am_eligible.sort(key=lambda w: (
            0 if w.id in baby_care_ids else 1,
            tallies[w.id]["AM"],
            random.random()
        ))
        am_chosen = am_eligible[:am_cap]
        for w in am_chosen:
            schedule[(w.id, d)] = "AM"
            tallies[w.id]["AM"] += 1
            unassigned_today = [u for u in unassigned_today if u.id != w.id]

        # Allocate SWING Shifts:
        swing_eligible = [w for w in unassigned_today if (w.id not in coming_from_off or w.id in baby_care_ids)]
        swing_eligible.sort(key=lambda w: (
            0 if w.id in baby_care_ids else 1,
            tallies[w.id]["SWING"],
            random.random()
        ))
        swing_chosen = swing_eligible[:swing_cap]
        for w in swing_chosen:
            schedule[(w.id, d)] = "SWING"
            tallies[w.id]["SWING"] += 1
            unassigned_today = [u for u in unassigned_today if u.id != w.id]

        # Rule 4: COMPLETE COVERAGE
        # "All staff shouuld be allocated a shift given any day, if they are not on shift they are on off, leave, sl or ph"
        # Remaining pool gets assigned OFF (never None / empty)
        for w in unassigned_today:
            schedule[(w.id, d)] = "OFF"
            tallies[w.id]["OFF"] += 1

    # Convert to database rows
    new_assignments = []
    for (wid, d), code in schedule.items():
        new_assignments.append({"worker_id": wid, "day": d, "shift_code": code})

    insert_assignments_db(new_assignments)
    return new_assignments

def get_schedule_grid_json():
    """Returns weekly schedule as a structured dictionary for JSON responses."""
    all_workers = get_all_workers_with_assignments()
    days_list = get_dynamic_days()
    grid = []
    for w in all_workers:
        assign_map = {a.day: (a.shift_code or "—") for a in w.assignments}
        row = {"worker_id": w.id, "worker_name": w.name, "days": {}}
        for d in days_list:
            row["days"][d] = assign_map.get(d, "—")
        grid.append(row)
    return grid

# ------------------------------------------------------------- OpenRouter AI Engine --
def call_openrouter(api_key, model, messages):
    """Executes a completion request to OpenRouter API with SSL & error handling."""
    if not api_key:
        return {"error": "OpenRouter API Key is missing. Please provide it in Settings or the AI Chat Drawer."}

    payload = {
        "model": model or "meta-llama/llama-3.3-70b-instruct:free",
        "messages": messages,
        "temperature": 0.2,
        "max_tokens": 4096,
    }
    data_bytes = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        "https://openrouter.ai/api/v1/chat/completions",
        data=data_bytes,
        headers={
            "Authorization": f"Bearer {api_key.strip()}",
            "Content-Type": "application/json",
            "HTTP-Referer": "http://127.0.0.1:5000",
            "X-Title": "RosterLine Shift Scheduler",
            "User-Agent": "RosterLine/1.0"
        }
    )
    ctx = ssl.create_default_context()
    try:
        with urllib.request.urlopen(req, context=ctx, timeout=60) as resp:
            body = resp.read().decode("utf-8")
            res_json = json.loads(body)
            choices = res_json.get("choices", [])
            if not choices:
                return {"error": "No response returned from the model."}
            msg_obj = choices[0].get("message", {})
            content = msg_obj.get("content") or msg_obj.get("reasoning") or ""
            return {"content": content, "raw": res_json}
    except urllib.error.HTTPError as e:
        err_msg = e.read().decode("utf-8", errors="ignore")
        try:
            err_json = json.loads(err_msg)
            msg = err_json.get("error", {}).get("message", err_msg)
        except Exception:
            msg = err_msg or str(e)
        return {"error": f"OpenRouter Error ({e.code}): {msg}"}
    except urllib.error.URLError as e:
        return {"error": f"Network Error connecting to OpenRouter: {e.reason}"}
    except Exception as e:
        return {"error": f"Unexpected error: {str(e)}"}

def extract_actions_from_response(content):
    """Extracts JSON actions from the LLM markdown response."""
    content = content or ""
    match = re.search(r"```(?:json)?\s*(\{[\s\S]*?\})\s*```", content)
    json_str = None
    if match:
        json_str = match.group(1)
    else:
        match_raw = re.search(r"(\{\s*\"actions\"\s*:\s*\[[\s\S]*?\]\s*\})", content)
        if match_raw:
            json_str = match_raw.group(1)

    clean_text = content
    if json_str:
        try:
            parsed = json.loads(json_str)
            actions = parsed.get("actions", [])
            clean_text = re.sub(r"```(?:json)?\s*\{[\s\S]*?\}\s*```", "", content).strip()
            return actions, clean_text
        except Exception:
            pass
    return [], clean_text

def build_system_prompt(workers, configs, days_list, start_str, end_str):
    """Constructs the comprehensive system prompt for the OpenRouter AI."""
    configs_summary = ", ".join([f"{c.code} ({c.label}, {c.start_time}-{c.end_time}, capacity: {c.capacity})" for c in configs])
    
    workers_text_list = []
    for w in workers:
        leaves = worker_leave_set(w)
        leave_str = f" [Approved Leave: {', '.join(leaves)}]" if leaves else ""
        baby_str = " [RESTRICTION: YOUNG BABY - ONLY AM/SWING, NEVER PM]" if is_baby_care_worker(w.name) else ""
        workers_text_list.append(f"- ID {w.id}: {w.name}{baby_str}{leave_str}")
    workers_text = "\n".join(workers_text_list)

    grid_lines = ["Worker ID | Name | " + " | ".join([format_day_short(d) for d in days_list])]
    for w in workers:
        assign_map = {a.day: (a.shift_code or "—") for a in w.assignments}
        row = [f"{w.id}", f"{w.name}"]
        for d in days_list:
            row.append(assign_map.get(d, "—"))
        grid_lines.append(" | ".join(row))
    current_grid_text = "\n".join(grid_lines)

    prompt = f"""You are RosterLine AI, an expert hospital and clinical shift scheduler.
You help allocate, modify, balance, and review shift rosters for the team.

CURRENT ROSTER CONTEXT:
- Date Range: {start_str or 'Current Week'} to {end_str or 'Current Week'}
- Dynamic Dates: {', '.join(days_list)}
- Configured Shifts: {configs_summary}
- Available Shift & Status Codes:
  * Working Shifts: AM, SWING (or SW), PM
  * Non-Working Statuses: OFF (Day Off), LEAVE (Approved Leave), SL (Sick Leave), PH (Public Holiday), SO (Sick Off), ST (Study Time)
- Team Members:
{workers_text}

CURRENT SCHEDULE TABLE:
{current_grid_text}

STRICT OPERATIONAL RULES (MANDATORY):
1. OFF-TO-PM REST TRANSITION:
   When any staff member is coming from their OFF day (or consecutive off days), if they work on the following day, they ONLY enter PM shifts.
   They must NEVER enter an AM or SWING shift immediately following an OFF day.

2. YOUNG BABIES RESTRICTION (Jacky & Eunice):
   Jacky and Eunice (spelled Equnice) have young babies at home.
   When working, they can ONLY be scheduled on AM and SWING (SW) shifts.
   They must NEVER be scheduled on a PM shift under any circumstances.

3. BALANCED AM & PM SHIFTS:
   Try to balance the PM and AM shifts across all eligible staff members over the schedule period.
   Avoid giving one worker too many PMs or AMs while others have none.

4. COMPLETE COVERAGE (EVERY STAFF GETS A STATUS EVERY DAY):
   All staff must be allocated a status on every single day.
   If a staff member is not on shift (working AM, SWING, or PM), they must be marked OFF, LEAVE, SL (sick leave), or PH (public holiday).
   No worker should EVER be left unassigned, blank, or '—'.

5. PUBLIC HOLIDAYS & SICK LEAVE:
   - Use 'PH' when designating a Public Holiday.
   - Use 'SL' when designating Sick Leave.

RESPONSE INSTRUCTIONS:
1. Explain your actions or answer the user clearly and concisely in markdown prose.
2. When you create, modify, swap, or re-allocate any shifts, you MUST append a JSON block at the end of your response with the exact actions to apply in this schema:
```json
{{
  "actions": [
    {{"worker_id": 1, "day": "2026-10-31", "shift_code": "PM"}},
    {{"worker_id": 2, "day": "2026-10-31", "shift_code": "AM"}}
  ]
}}
```
Only include the JSON block if shifts should actually be updated in the database. If the user asks a question, requests a summary, or needs advice without modifying data, omit the JSON actions block.
"""
    return prompt

def apply_chat_actions(actions):
    """Saves each shift action returned by the LLM."""
    count = 0
    valid_codes = {"AM", "SWING", "SW", "PM", "OFF", "LEAVE", "SL", "PH", "SO", "ST"}
    for act in actions:
        worker_id = act.get("worker_id")
        day = act.get("day")
        code = act.get("shift_code")
        if worker_id and day:
            try:
                worker_id = int(worker_id)
                if code:
                    code = str(code).upper().strip()
                    if code in ("SW", "SWING_SHIFT"):
                        code = "SWING"
                    if code not in valid_codes:
                        code = None
                else:
                    code = None
                assign_single_shift_db(worker_id, day, code)
                count += 1
            except Exception as e:
                print(f"Error applying chat action {act}: {e}")
    return count

# ----------------------------------------------------------------- views --
@app.route("/")
def index():
    configs = get_shift_configs()

    view_mode = request.args.get("view", "weekly")
    if view_mode not in ("daily", "weekly"):
        view_mode = "weekly"

    days_list = get_dynamic_days()
    days_formatted = {d: format_day_short(d) for d in days_list}
    
    current_day = request.args.get("day")
    if current_day not in days_list:
        current_day = days_list[0] if days_list else date.today().strftime("%Y-%m-%d")

    groups, unassigned, off_workers, leave_workers, sl_workers, ph_workers, all_workers = grouped_workers_for_day(current_day)

    # Build weekly grid data
    weekly_schedule = []
    for w in all_workers:
        assign_map = {a.day: (a.shift_code or "—") for a in w.assignments}
        leaves = worker_leave_set(w)
        row = {"worker": w, "days": {}}
        for d in days_list:
            if d in assign_map:
                row["days"][d] = assign_map[d]
            else:
                try:
                    weekday_name = datetime.strptime(d, "%Y-%m-%d").strftime("%A")
                except:
                    weekday_name = d
                if weekday_name in leaves:
                    row["days"][d] = "LEAVE"
                else:
                    row["days"][d] = "—"
        weekly_schedule.append(row)

    openrouter_key = get_setting("openrouter_api_key", "")
    openrouter_model = get_setting("openrouter_model", "meta-llama/llama-3.3-70b-instruct:free")

    return render_template(
        "index.html",
        configs=configs,
        groups=groups,
        unassigned=unassigned,
        off_workers=off_workers,
        leave_workers=leave_workers,
        sl_workers=sl_workers,
        ph_workers=ph_workers,
        shift_meta=shift_meta_dict(),
        total=len(all_workers),
        today=date.today().strftime("%A, %d %B %Y"),
        today_iso=date.today().strftime("%Y-%m-%d"),
        days=days_list,
        days_formatted=days_formatted,
        current_day=current_day,
        view_mode=view_mode,
        weekly_schedule=weekly_schedule,
        has_openrouter_key=bool(openrouter_key),
        openrouter_model=openrouter_model,
    )


@app.route("/allocate", methods=["POST"])
def allocate():
    workers = get_all_workers_with_assignments()
    if not workers:
        flash("Add some workers first, then allocate shifts.", "warning")
        return redirect(url_for("index"))

    run_auto_allocation()
    flash("Shifts auto-allocated adhering to baby-care restrictions (AM/SWING only), post-off PM transitions, balanced AM/PM shifts, and complete coverage.", "success")

    view = request.form.get("view", "weekly")
    days_list = get_dynamic_days()
    cday = request.form.get("day", days_list[0] if days_list else date.today().strftime("%Y-%m-%d"))
    return redirect(url_for("index", view=view, day=cday))


@app.route("/reset", methods=["POST"])
def reset():
    days_list = get_dynamic_days()
    if days_list:
        delete_assignments_for_days_db(days_list)
    flash("Shift assignments for the current date range cleared.", "info")
    view = request.form.get("view", "weekly")
    cday = request.form.get("day", days_list[0] if days_list else date.today().strftime("%Y-%m-%d"))
    return redirect(url_for("index", view=view, day=cday))


@app.route("/assign", methods=["POST"])
def assign_shift():
    """AJAX endpoint: update a single worker/day shift assignment."""
    data = request.get_json(force=True)
    worker_id  = data.get("worker_id")
    day        = data.get("day")
    shift_code = data.get("shift_code")

    try:
        worker_id = int(worker_id)
    except (TypeError, ValueError):
        return jsonify({"ok": False, "error": "Invalid worker_id."}), 400

    if not worker_id or not day:
        return jsonify({"ok": False, "error": "Invalid parameters."}), 400

    success = assign_single_shift_db(worker_id, day, shift_code)
    if success:
        return jsonify({"ok": True})
    return jsonify({"ok": False, "error": "Database error saving shift."}), 500


# ------------------------------------------------------------- AI Chatbot Endpoints --
@app.route("/api/chat", methods=["POST"])
def api_chat():
    data = request.get_json(force=True) or {}
    user_message = data.get("message", "").strip()
    history = data.get("history", [])
    model_override = data.get("model", "").strip()
    api_key_override = data.get("api_key", "").strip()

    if not user_message:
        return jsonify({"ok": False, "error": "Please enter a message."}), 400

    api_key = api_key_override or get_setting("openrouter_api_key") or os.environ.get("OPENROUTER_API_KEY", "").strip()
    model = model_override or get_setting("openrouter_model") or os.environ.get("OPENROUTER_MODEL", "meta-llama/llama-3.3-70b-instruct:free").strip()

    workers = get_all_workers_with_assignments()
    configs = get_shift_configs()
    days_list = get_dynamic_days()
    start_str, end_str = get_roster_date_range()

    system_prompt = build_system_prompt(workers, configs, days_list, start_str, end_str)

    # If no API key configured, check if user requested allocation or rule check
    if not api_key:
        lower_msg = user_message.lower()
        if any(k in lower_msg for k in ["allocate", "generate", "auto-allocate", "schedule", "assign"]):
            new_assignments = run_auto_allocation()
            return jsonify({
                "ok": True,
                "reply": "I have allocated the shifts using the defined system prompt constraints:\n\n"
                         "1. **Post-OFF Transition**: Staff coming from an OFF day are assigned PM shifts when working.\n"
                         "2. **Young Babies Restriction**: Jacky and Eunice are restricted to AM and SWING shifts only.\n"
                         "3. **Balanced Workload**: AM and PM shifts are balanced across all eligible staff.\n"
                         "4. **Full Coverage**: Every staff member is allocated a shift or non-working status (OFF, LEAVE, SL, PH) on every single day.\n\n"
                         "*(Tip: To chat conversationally and make custom edits with OpenRouter, add your free API key in the ⚙️ settings panel above.)*",
                "actions_count": len(new_assignments),
                "model": "Built-in Solver",
                "updated_table": get_schedule_grid_json()
            })
        else:
            return jsonify({
                "ok": True,
                "reply": "To chat conversationally with the AI, please enter your **OpenRouter API Key** by clicking the ⚙️ icon in the header above or visiting **Settings**.\n\n"
                         "You can get a 100% free API key at [OpenRouter.ai](https://openrouter.ai/keys).\n\n"
                         "Alternatively, you can click **'⚡ Auto-allocate'** below to generate the roster with all 4 strict rules right now!",
                "needs_key": True
            })

    # Prepare message payload
    messages = [{"role": "system", "content": system_prompt}]
    for h in history[-8:]:
        role = h.get("role")
        content = h.get("content")
        if role in ("user", "assistant") and content:
            messages.append({"role": role, "content": content})
    messages.append({"role": "user", "content": user_message})

    resp = call_openrouter(api_key, model, messages)
    if "error" in resp:
        return jsonify({"ok": False, "error": resp["error"]}), 500

    content = resp.get("content", "")
    actions, clean_reply = extract_actions_from_response(content)

    applied_count = 0
    if actions:
        applied_count = apply_chat_actions(actions)

    return jsonify({
        "ok": True,
        "reply": clean_reply,
        "actions_count": applied_count,
        "actions": actions,
        "model": model,
        "updated_table": get_schedule_grid_json() if applied_count > 0 else None
    })

@app.route("/api/openrouter-config", methods=["GET", "POST"])
def api_openrouter_config():
    if request.method == "POST":
        data = request.get_json(force=True) or {}
        key = data.get("api_key")
        model = data.get("model")
        if key is not None:
            set_setting("openrouter_api_key", key.strip())
        if model:
            set_setting("openrouter_model", model.strip())
        return jsonify({"ok": True})
    
    # GET
    key = get_setting("openrouter_api_key", "")
    model = get_setting("openrouter_model", "meta-llama/llama-3.3-70b-instruct:free")
    key_preview = f"{key[:8]}...{key[-4:]}" if len(key) > 12 else ("Set" if key else "")
    return jsonify({
        "has_key": bool(key),
        "key_preview": key_preview,
        "model": model
    })

@app.route("/api/roster-data")
def api_roster_data():
    return jsonify({
        "days": get_dynamic_days(),
        "schedule": get_schedule_grid_json()
    })

# -------------------------------------------------------- settings page --
@app.route("/settings")
def settings():
    configs = get_shift_configs()
    workers = get_all_workers_with_assignments()
    start_date, end_date = get_roster_date_range()
    openrouter_key = get_setting("openrouter_api_key", "")
    openrouter_model = get_setting("openrouter_model", "meta-llama/llama-3.3-70b-instruct:free")
    
    return render_template(
        "settings.html",
        configs=configs,
        workers=workers,
        start_date=start_date,
        end_date=end_date,
        days=["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"],
        openrouter_key=openrouter_key,
        openrouter_model=openrouter_model
    )

@app.route("/settings/save", methods=["POST"])
def settings_save():
    configs = get_shift_configs()
    errors = []
    for cfg in configs:
        prefix     = f"shift_{cfg.code}_"
        label      = request.form.get(prefix + "label", "").strip()
        start_time = request.form.get(prefix + "start_time", "").strip()
        end_time   = request.form.get(prefix + "end_time", "").strip()
        cap_str    = request.form.get(prefix + "capacity", "1").strip()

        if not label:
            errors.append(f"Label for {cfg.code} cannot be empty.")
            continue
        try:
            cap = int(cap_str)
            if cap < 1:
                raise ValueError
        except ValueError:
            errors.append(f"Capacity for {cfg.code} must be a positive integer.")
            continue

        update_shift_config_db(cfg.id, label, start_time, end_time, cap)

    # Roster Date Range Save
    roster_start = request.form.get("roster_start", "").strip()
    roster_end = request.form.get("roster_end", "").strip()
    if roster_start and roster_end:
        set_roster_date_range(roster_start, roster_end)

    # OpenRouter API Key & Model Save
    openrouter_key = request.form.get("openrouter_api_key")
    if openrouter_key is not None:
        set_setting("openrouter_api_key", openrouter_key.strip())

    openrouter_model = request.form.get("openrouter_model")
    if openrouter_model:
        set_setting("openrouter_model", openrouter_model.strip())

    if errors:
        for e in errors:
            flash(e, "warning")
    else:
        flash("Settings saved successfully.", "success")

    return redirect(url_for("settings"))

@app.route("/settings/leave", methods=["POST"])
def settings_leave():
    workers = get_all_workers_with_assignments()
    WEEKDAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    for w in workers:
        chosen = [d for d in WEEKDAYS if request.form.get(f"leave_{w.id}_{d}")]
        leave_str = ",".join(chosen)
        update_worker_leave_db(w.id, leave_str)
            
    flash("Leave settings saved successfully.", "success")
    return redirect(url_for("settings"))

# ------------------------------------------------------- workers page --
@app.route("/workers")
def workers():
    all_workers = get_all_workers_with_assignments()
    groups, _, _, _, _, _, _ = grouped_workers_for_day(date.today().strftime("%Y-%m-%d"))
    return render_template(
        "workers.html",
        all_workers=all_workers,
        shift_meta=shift_meta_dict(),
    )

@app.route("/workers/add", methods=["POST"])
def add_worker():
    name = request.form.get("name", "").strip()
    if not name:
        flash("Enter a name before adding a worker.", "warning")
    elif name_exists(name):
        flash(f'"{name}" is already on the list.', "warning")
    else:
        add_worker_db(name)
        flash(f'"{name}" added.', "success")
    return redirect(url_for("workers"))

@app.route("/workers/delete/<int:worker_id>", methods=["POST"])
def delete_worker(worker_id):
    name = delete_worker_db(worker_id)
    flash(f'"{name}" removed.', "info")
    return redirect(url_for("workers"))

@app.route("/workers/upload", methods=["POST"])
def upload_workers():
    file = request.files.get("file")
    if not file or file.filename == "":
        flash("Choose an Excel file to upload.", "warning")
        return redirect(url_for("workers"))

    if not file.filename.lower().endswith((".xlsx", ".xlsm")):
        flash("Please upload a .xlsx file.", "warning")
        return redirect(url_for("workers"))

    try:
        wb = load_workbook(file, read_only=True, data_only=True)
        sheet = wb.active
    except Exception:
        flash("Could not read that file. Make sure it's a valid Excel workbook.", "warning")
        return redirect(url_for("workers"))

    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        flash("That sheet looks empty.", "warning")
        return redirect(url_for("workers"))

    header = [str(c).strip().lower() if c is not None else "" for c in rows[0]]
    name_col = header.index("name") if "name" in header else 0
    data_rows = rows[1:] if "name" in header else rows

    added, skipped = 0, 0
    seen = set()
    new_workers = []
    
    for row in data_rows:
        if not row or name_col >= len(row):
            continue
        raw = row[name_col]
        if raw is None:
            continue
        name = str(raw).strip()
        if not name:
            continue
        key = name.lower()
        if key in seen or name_exists(name):
            skipped += 1
            continue
        seen.add(key)
        new_workers.append({"name": name})
        added += 1

    if new_workers:
        if supabase:
            try:
                supabase.table("worker").insert(new_workers).execute()
            except Exception as e:
                print(f"Supabase bulk upload error: {e}")
        conn = get_db_conn()
        for nw in new_workers:
            conn.execute("INSERT INTO worker (name, leave_days) VALUES (?, '')", (nw["name"],))
        conn.commit()
        conn.close()

    msg = f"Imported {added} worker(s)."
    if skipped:
        msg += f" Skipped {skipped} duplicate/blank row(s)."
    flash(msg, "success" if added else "warning")
    return redirect(url_for("workers"))

@app.route("/template")
def download_template():
    wb    = Workbook()
    sheet = wb.active
    sheet.title = "Workers"
    sheet["A1"] = "Name"
    sheet["A1"].font = Font(name="Arial", bold=True, color="FFFFFF")
    sheet["A1"].fill = PatternFill("solid", fgColor="2B2F4C")
    sheet["A2"] = "Jane Doe"
    sheet["A2"].font = Font(name="Arial", italic=True, color="8B90A8")
    sheet.column_dimensions["A"].width = 28
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="worker_import_template.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/export")
def export_schedule():
    all_workers = get_all_workers_with_assignments()

    wb = Workbook()
    sheet = wb.active
    sheet.title = "Weekly Schedule"

    FILL = {
        "AM":    PatternFill("solid", fgColor="F2B872"),
        "SWING": PatternFill("solid", fgColor="E8734A"),
        "PM":    PatternFill("solid", fgColor="5C6BC0"),
        "OFF":   PatternFill("solid", fgColor="E2E4E9"),
        "LEAVE": PatternFill("solid", fgColor="A6ADB5"),
        "SO":    PatternFill("solid", fgColor="E84A6A"),
        "ST":    PatternFill("solid", fgColor="4AE870"),
        "SL":    PatternFill("solid", fgColor="FCA5A5"),
        "PH":    PatternFill("solid", fgColor="C7D2FE"),
    }
    DARK_TEXT = {"PM", "LEAVE", "SO", "SL"}

    title_font = Font(name="Arial", bold=True, size=14, color="1B2036")
    hdr_font   = Font(name="Arial", bold=True, size=11, color="FFFFFF")
    hdr_fill   = PatternFill("solid", fgColor="2B2F4C")
    body_font  = Font(name="Arial", size=10)
    white_font = Font(name="Arial", size=10, color="FFFFFF")
    center     = Alignment(horizontal="center", vertical="center")
    left       = Alignment(horizontal="left",   vertical="center")
    thin       = Side(style="thin", color="D9D9D9")
    border     = Border(left=thin, right=thin, top=thin, bottom=thin)

    sheet.merge_cells("A1:H1")
    sheet["A1"] = f"Weekly Shift Schedule — {date.today().strftime('%d %B %Y')}"
    sheet["A1"].font = title_font
    sheet["A1"].alignment = left
    sheet.row_dimensions[1].height = 28

    days_list = get_dynamic_days()
    headers = ["Worker Name"] + [format_day_short(d) for d in days_list]
    sheet.row_dimensions[3].height = 24
    for ci, h in enumerate(headers, start=1):
        c = sheet.cell(row=3, column=ci, value=h)
        c.font = hdr_font; c.fill = hdr_fill
        c.alignment = center; c.border = border

    sheet.column_dimensions["A"].width = 26
    for col in ["B","C","D","E","F","G","H"]:
        sheet.column_dimensions[col].width = 14

    for ri, w in enumerate(all_workers, start=4):
        sheet.row_dimensions[ri].height = 20
        nc = sheet.cell(row=ri, column=1, value=w.name)
        nc.font = body_font; nc.alignment = left; nc.border = border

        assign_map = {a.day: (a.shift_code or "—") for a in w.assignments}
        leaves = worker_leave_set(w)

        for ci, day_str in enumerate(days_list, start=2):
            if day_str in assign_map:
                val = assign_map[day_str]
            else:
                try:
                    weekday_name = datetime.strptime(day_str, "%Y-%m-%d").strftime("%A")
                except:
                    weekday_name = day_str
                if weekday_name in leaves:
                    val = "LEAVE"
                else:
                    val = "—"

            cell = sheet.cell(row=ri, column=ci, value=val)
            cell.alignment = center; cell.border = border
            if val in FILL:
                cell.fill = FILL[val]
                cell.font = white_font if val in DARK_TEXT else body_font
            else:
                cell.font = body_font

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    fname = f"weekly_schedule_{date.today().isoformat()}.xlsx"
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    app.run(debug=True)
