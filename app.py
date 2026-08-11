import io
import os
import random
from datetime import datetime, date
from types import SimpleNamespace

from flask import (
    Flask, render_template, request, redirect, url_for, flash, send_file
)
from supabase import create_client, Client
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

app = Flask(__name__)

app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "change-this-secret-key-in-production")

SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_SERVICE_KEY")

supabase: Client = None
if SUPABASE_URL and SUPABASE_KEY:
    supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
else:
    print("WARNING: SUPABASE_URL and SUPABASE_SERVICE_KEY must be set!")

DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]

_DEFAULT_SHIFTS = [
    dict(code="AM",    label="AM Shift",    start_time="6:00 AM",  end_time="3:00 PM",  capacity=3, sort_order=0),
    dict(code="SWING", label="Swing Shift", start_time="11:00 AM", end_time="8:00 PM",  capacity=1, sort_order=1),
    dict(code="PM",    label="PM Shift",    start_time="3:00 PM",  end_time="12:00 AM", capacity=3, sort_order=2),
]

def init_db():
    if not supabase:
        return
    try:
        res = supabase.table("shift_config").select("id", count="exact").execute()
        if res.count == 0:
            supabase.table("shift_config").insert(_DEFAULT_SHIFTS).execute()
    except Exception as e:
        print(f"Error seeding shifts: {e}")

init_db()

# ------------------------------------------------------------- utilities --

def get_shift_configs():
    if not supabase: return []
    try:
        res = supabase.table("shift_config").select("*").order("sort_order").execute()
        return [SimpleNamespace(**d) for d in res.data]
    except Exception as e:
        print(f"Error fetching shift configs: {e}")
        return []

def shift_meta_dict():
    return {s.code: s for s in get_shift_configs()}

def worker_leave_set(worker):
    """Return set of leave day names for a worker."""
    if not getattr(worker, 'leave_days', None):
        return set()
    return {d.strip() for d in worker.leave_days.split(",") if d.strip()}

def get_all_workers_with_assignments():
    if not supabase: return []
    try:
        res = supabase.table("worker").select("*, assignments:assignment(*)").order("name").execute()
        workers = []
        for d in res.data:
            w = SimpleNamespace(**d)
            w.assignments = [SimpleNamespace(**a) for a in getattr(w, 'assignments', [])]
            workers.append(w)
        return workers
    except Exception as e:
        print(f"Error fetching workers: {e}")
        return []


def grouped_workers_for_day(day):
    """Return (groups, unassigned, off_workers, leave_workers, all_workers) for a given weekday."""
    configs    = get_shift_configs()
    codes      = [c.code for c in configs]
    all_workers = get_all_workers_with_assignments()

    groups        = {code: [] for code in codes}
    unassigned    = []
    off_workers   = []
    leave_workers = []

    for w in all_workers:
        assign = next((a for a in w.assignments if a.day == day), None)
        if assign:
            if assign.shift_code in groups:
                groups[assign.shift_code].append(w)
            elif assign.shift_code == "OFF":
                off_workers.append(w)
            elif assign.shift_code == "LEAVE":
                leave_workers.append(w)
            else:
                unassigned.append(w)
        else:
            # No assignment yet — classify by leave settings
            if day in worker_leave_set(w):
                leave_workers.append(w)
            else:
                unassigned.append(w)

    return groups, unassigned, off_workers, leave_workers, all_workers

def name_exists(name):
    if not supabase: return False
    res = supabase.table("worker").select("id").ilike("name", name.strip()).execute()
    return len(res.data) > 0


# ----------------------------------------------------------------- views --

@app.route("/")
def index():
    configs = get_shift_configs()

    view_mode = request.args.get("view", "daily")
    if view_mode not in ("daily", "weekly"):
        view_mode = "daily"

    current_day = request.args.get("day", date.today().strftime("%A"))
    if current_day not in DAYS:
        current_day = DAYS[0]

    groups, unassigned, off_workers, leave_workers, all_workers = grouped_workers_for_day(current_day)

    # Build weekly grid data
    weekly_schedule = []
    for w in all_workers:
        assign_map = {a.day: (a.shift_code or "—") for a in w.assignments}
        leaves = worker_leave_set(w)
        row = {"worker": w, "days": {}}
        for d in DAYS:
            if d in assign_map:
                row["days"][d] = assign_map[d]
            elif d in leaves:
                row["days"][d] = "LEAVE"
            else:
                row["days"][d] = "—"
        weekly_schedule.append(row)

    return render_template(
        "index.html",
        configs=configs,
        groups=groups,
        unassigned=unassigned,
        off_workers=off_workers,
        leave_workers=leave_workers,
        shift_meta=shift_meta_dict(),
        total=len(all_workers),
        today=date.today().strftime("%A, %d %B %Y"),
        today_weekday=date.today().strftime("%A"),
        days=DAYS,
        current_day=current_day,
        view_mode=view_mode,
        weekly_schedule=weekly_schedule,
    )


@app.route("/allocate", methods=["POST"])
def allocate():
    if not supabase:
        flash("Supabase not configured.", "danger")
        return redirect(url_for("index"))
        
    workers = get_all_workers_with_assignments()
    if not workers:
        flash("Add some workers first, then allocate shifts.", "warning")
        return redirect(url_for("index"))

    configs = get_shift_configs()

    # Wipe existing weekly assignments
    supabase.table("assignment").delete().neq("id", -1).execute()

    # Build per-day availability counts (excluding leave days)
    day_available = {}
    for d in DAYS:
        day_available[d] = sum(1 for w in workers if d not in worker_leave_set(w))

    # Shuffle for fairness
    workers_list = list(workers)
    random.shuffle(workers_list)

    off_day_map = {}   # worker_id -> chosen OFF day
    
    new_assignments = []

    # 1. Assign LEAVE rows and choose each worker's 1 day off
    for w in workers_list:
        leaves = worker_leave_set(w)
        for d in leaves:
            new_assignments.append({"worker_id": w.id, "day": d, "shift_code": "LEAVE"})

        available = [d for d in DAYS if d not in leaves]
        if not available:
            continue   # fully on leave

        # Pick the day with the most available workers (most room to absorb the absence)
        off_day = max(available, key=lambda d: day_available[d])
        off_day_map[w.id] = off_day
        day_available[off_day] -= 1
        new_assignments.append({"worker_id": w.id, "day": off_day, "shift_code": "OFF"})

    # 2. Day-by-day shift allocation
    for d in DAYS:
        active = [
            w for w in workers
            if d not in worker_leave_set(w) and off_day_map.get(w.id) != d
        ]
        random.shuffle(active)

        idx = 0
        for cfg in configs:
            for _ in range(cfg.capacity):
                if idx < len(active):
                    new_assignments.append({"worker_id": active[idx].id, "day": d, "shift_code": cfg.code})
                    idx += 1
        # Remaining actives are "idle/unassigned" that day
        while idx < len(active):
            new_assignments.append({"worker_id": active[idx].id, "day": d, "shift_code": None})
            idx += 1

    if new_assignments:
        supabase.table("assignment").insert(new_assignments).execute()

    flash("Weekly shifts auto-allocated — every worker has 1 day off.", "success")

    view  = request.form.get("view", "daily")
    cday  = request.form.get("day", date.today().strftime("%A"))
    return redirect(url_for("index", view=view, day=cday))


@app.route("/reset", methods=["POST"])
def reset():
    if supabase:
        supabase.table("assignment").delete().neq("id", -1).execute()
    flash("All weekly shift assignments cleared.", "info")
    view = request.form.get("view", "daily")
    cday = request.form.get("day", date.today().strftime("%A"))
    return redirect(url_for("index", view=view, day=cday))


# -------------------------------------------------------- settings page --

@app.route("/settings")
def settings():
    configs = get_shift_configs()
    workers = get_all_workers_with_assignments()
    return render_template("settings.html", configs=configs, workers=workers, days=DAYS)


@app.route("/settings/save", methods=["POST"])
def settings_save():
    configs = get_shift_configs()
    errors  = []
    for cfg in configs:
        prefix     = f"shift_{cfg.code}_"
        label      = request.form.get(prefix + "label",      "").strip()
        start_time = request.form.get(prefix + "start_time", "").strip()
        end_time   = request.form.get(prefix + "end_time",   "").strip()
        cap_str    = request.form.get(prefix + "capacity",   "1").strip()

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
            
        if supabase:
            supabase.table("shift_config").update({
                "label": label,
                "start_time": start_time,
                "end_time": end_time,
                "capacity": cap
            }).eq("id", cfg.id).execute()

    if errors:
        for e in errors:
            flash(e, "warning")
    else:
        flash("Shift settings saved successfully.", "success")

    return redirect(url_for("settings"))


@app.route("/settings/leave", methods=["POST"])
def settings_leave():
    workers = get_all_workers_with_assignments()
    for w in workers:
        chosen = [d for d in DAYS if request.form.get(f"leave_{w.id}_{d}")]
        leave_str = ",".join(chosen)
        if supabase:
            supabase.table("worker").update({"leave_days": leave_str}).eq("id", w.id).execute()
            
    flash("Leave settings saved successfully.", "success")
    return redirect(url_for("settings"))


# ------------------------------------------------------- workers page --

@app.route("/workers")
def workers():
    all_workers = get_all_workers_with_assignments()
    # For the badge column use Monday as a reference day
    groups, _, _, _, _ = grouped_workers_for_day("Monday")
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
        if supabase:
            supabase.table("worker").insert({"name": name}).execute()
        flash(f'"{name}" added.', "success")
    return redirect(url_for("workers"))


@app.route("/workers/delete/<int:worker_id>", methods=["POST"])
def delete_worker(worker_id):
    if supabase:
        # Get worker to show name in flash
        res = supabase.table("worker").select("name").eq("id", worker_id).execute()
        if res.data:
            name = res.data[0]["name"]
            supabase.table("worker").delete().eq("id", worker_id).execute()
            flash(f'"{name}" removed.', "info")
        else:
            flash("Worker not found.", "warning")
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
        wb    = load_workbook(file, read_only=True, data_only=True)
        sheet = wb.active
    except Exception:
        flash("Could not read that file. Make sure it's a valid Excel workbook.", "warning")
        return redirect(url_for("workers"))

    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        flash("That sheet looks empty.", "warning")
        return redirect(url_for("workers"))

    header   = [str(c).strip().lower() if c is not None else "" for c in rows[0]]
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

    if new_workers and supabase:
        supabase.table("worker").insert(new_workers).execute()

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

    wb    = Workbook()
    sheet = wb.active
    sheet.title = "Weekly Schedule"

    FILL = {
        "AM":    PatternFill("solid", fgColor="F2B872"),
        "SWING": PatternFill("solid", fgColor="E8734A"),
        "PM":    PatternFill("solid", fgColor="5C6BC0"),
        "OFF":   PatternFill("solid", fgColor="E2E4E9"),
        "LEAVE": PatternFill("solid", fgColor="A6ADB5"),
    }
    DARK_TEXT = {"PM", "LEAVE"}

    title_font  = Font(name="Arial", bold=True, size=14, color="1B2036")
    hdr_font    = Font(name="Arial", bold=True, size=11, color="FFFFFF")
    hdr_fill    = PatternFill("solid", fgColor="2B2F4C")
    body_font   = Font(name="Arial", size=10)
    white_font  = Font(name="Arial", size=10, color="FFFFFF")
    center      = Alignment(horizontal="center", vertical="center")
    left        = Alignment(horizontal="left",   vertical="center")
    thin        = Side(style="thin", color="D9D9D9")
    border      = Border(left=thin, right=thin, top=thin, bottom=thin)

    sheet.merge_cells("A1:H1")
    sheet["A1"] = f"Weekly Shift Schedule — {date.today().strftime('%d %B %Y')}"
    sheet["A1"].font = title_font
    sheet["A1"].alignment = left
    sheet.row_dimensions[1].height = 28

    headers = ["Worker Name"] + DAYS
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
        leaves     = worker_leave_set(w)

        for ci, day in enumerate(DAYS, start=2):
            if day in assign_map:
                val = assign_map[day]
            elif day in leaves:
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
