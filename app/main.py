from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field
from typing import List, Optional
from pathlib import Path
from datetime import datetime
from io import BytesIO
import re
import sqlite3
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)
DB_PATH = DATA_DIR / "urenregistratie.db"
LEGACY_WORKBOOK = DATA_DIR / "urenregistratie.xlsx"

app = FastAPI(title="UrenApp")
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "app" / "static")), name="static")
DAYS = ["Maandag", "Dinsdag", "Woensdag", "Donderdag", "Vrijdag", "Zaterdag", "Zondag"]

class WorkLine(BaseModel):
    day: str
    job: str = Field(min_length=1, max_length=120)
    hours: float = Field(ge=0, le=24)
    free_hours: float = Field(default=0, ge=0, le=24)
    note: Optional[str] = Field(default="", max_length=250)

class Submission(BaseModel):
    name: str = Field(min_length=2, max_length=60)
    week: int = Field(ge=1, le=53)
    zero_hours_contract: bool = False
    lines: List[WorkLine]

def period_for_week(week: int) -> int:
    return ((week - 1) // 4) + 1

def safe_sheet_name(name: str) -> str:
    name = re.sub(r"[\\/*?:\[\]]", "-", name).strip()
    return (name or "Medewerker")[:31]

def db():
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = db()
    conn.executescript("""
    PRAGMA journal_mode=WAL;
    CREATE TABLE IF NOT EXISTS submissions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        name_key TEXT NOT NULL,
        week INTEGER NOT NULL,
        period INTEGER NOT NULL,
        zero_hours_contract INTEGER NOT NULL DEFAULT 0,
        submitted_at TEXT NOT NULL,
        UNIQUE(name_key, week)
    );
    CREATE TABLE IF NOT EXISTS work_lines (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        submission_id INTEGER NOT NULL,
        day TEXT NOT NULL,
        job TEXT NOT NULL,
        work_hours REAL NOT NULL DEFAULT 0,
        free_hours REAL NOT NULL DEFAULT 0,
        note TEXT NOT NULL DEFAULT '',
        FOREIGN KEY(submission_id) REFERENCES submissions(id) ON DELETE CASCADE
    );
    """)
    conn.commit()
    conn.close()
    migrate_legacy_excel_if_needed()

def migrate_legacy_excel_if_needed():
    if not LEGACY_WORKBOOK.exists():
        return
    conn = db()
    count = conn.execute("SELECT COUNT(*) FROM submissions").fetchone()[0]
    if count > 0:
        conn.close()
        return
    try:
        wb = load_workbook(LEGACY_WORKBOOK, data_only=True)
        ignored = {"Overzicht", "Week-overzicht", "Periode-overzicht"}
        for sheet_name in wb.sheetnames:
            if sheet_name in ignored:
                continue
            ws = wb[sheet_name]
            grouped = {}
            for r in range(2, ws.max_row + 1):
                week = ws.cell(r, 2).value
                day = ws.cell(r, 3).value
                job = ws.cell(r, 4).value
                if not week or not day or not job:
                    continue
                key = int(week)
                grouped.setdefault(key, []).append({
                    "day": str(day),
                    "job": str(job),
                    "work": float(ws.cell(r, 5).value or 0),
                    "free": float(ws.cell(r, 6).value or 0),
                    "note": str(ws.cell(r, 7).value or ""),
                    "submitted_at": str(ws.cell(r, 8).value or datetime.now().strftime("%d-%m-%Y %H:%M")),
                })
            for week, lines in grouped.items():
                name = sheet_name
                name_key = name.strip().lower()
                submitted_at = lines[-1]["submitted_at"]
                period = period_for_week(week)
                zero = 1 if any(x["free"] > 0 for x in lines) else 0
                cur = conn.execute(
                    "INSERT OR IGNORE INTO submissions(name,name_key,week,period,zero_hours_contract,submitted_at) VALUES(?,?,?,?,?,?)",
                    (name, name_key, week, period, zero, submitted_at),
                )
                if cur.lastrowid:
                    sid = cur.lastrowid
                    for x in lines:
                        conn.execute(
                            "INSERT INTO work_lines(submission_id,day,job,work_hours,free_hours,note) VALUES(?,?,?,?,?,?)",
                            (sid, x["day"], x["job"], x["work"], x["free"], x["note"]),
                        )
        conn.commit()
    finally:
        conn.close()

def save_submission(sub: Submission):
    name = sub.name.strip()
    name_key = name.lower()
    period = period_for_week(sub.week)
    submitted_at = datetime.now().strftime("%d-%m-%Y %H:%M")
    valid_lines = []
    for line in sub.lines:
        if line.hours == 0 and line.free_hours == 0:
            continue
        valid_lines.append(line)
    if not valid_lines:
        raise HTTPException(400, "Vul minimaal 1 urenregel in.")

    conn = db()
    try:
        conn.execute("BEGIN IMMEDIATE")
        old = conn.execute("SELECT id FROM submissions WHERE name_key=? AND week=?", (name_key, sub.week)).fetchone()
        if old:
            conn.execute("DELETE FROM work_lines WHERE submission_id=?", (old["id"],))
            conn.execute("DELETE FROM submissions WHERE id=?", (old["id"],))
        cur = conn.execute(
            "INSERT INTO submissions(name,name_key,week,period,zero_hours_contract,submitted_at) VALUES(?,?,?,?,?,?)",
            (name, name_key, sub.week, period, 1 if sub.zero_hours_contract else 0, submitted_at),
        )
        sid = cur.lastrowid
        work_total = 0.0
        free_total = 0.0
        for line in valid_lines:
            work = float(line.hours)
            free = float(line.free_hours) if sub.zero_hours_contract else 0.0
            work_total += work
            free_total += free
            conn.execute(
                "INSERT INTO work_lines(submission_id,day,job,work_hours,free_hours,note) VALUES(?,?,?,?,?,?)",
                (sid, line.day, line.job.strip(), work, free, (line.note or "").strip()),
            )
        conn.commit()
        return {"period": period, "work_total": round(work_total, 2), "free_total": round(free_total, 2)}
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def style_header(ws, row=1):
    fill = PatternFill("solid", fgColor="1F4E78")
    for c in ws[row]:
        c.font = Font(color="FFFFFF", bold=True)
        c.fill = fill
        c.alignment = Alignment(vertical="center")
    ws.row_dimensions[row].height = 22

def autosize(ws):
    widths = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                widths[cell.column] = max(widths.get(cell.column, 0), len(str(cell.value)))
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = min(max(width + 2, 10), 36)

def fetch_summary_rows():
    conn = db()
    rows = conn.execute("""
        SELECT s.id, s.submitted_at, s.name, s.week, s.period, s.zero_hours_contract,
               COALESCE(SUM(w.work_hours),0) AS work_hours,
               COALESCE(SUM(w.free_hours),0) AS free_hours
        FROM submissions s
        LEFT JOIN work_lines w ON w.submission_id=s.id
        GROUP BY s.id
        ORDER BY s.week DESC, s.name COLLATE NOCASE ASC
    """).fetchall()
    conn.close()
    return rows

def build_excel():
    wb = Workbook()
    ov = wb.active
    ov.title = "Overzicht"
    ov.append(["Ingediend op", "Naam", "Periode", "Week", "Werkuren", "Vrije uren", "Totaal uren", "0-urencontract"])
    style_header(ov); ov.freeze_panes = "A2"

    wo = wb.create_sheet("Week-overzicht")
    wo.append(["Week", "Periode", "Naam", "Werkuren", "Vrije uren", "Totaal uren", "Ingediend op"])
    style_header(wo); wo.freeze_panes = "A2"

    po = wb.create_sheet("Periode-overzicht")
    po.append(["Periode", "Week", "Naam", "Werkuren", "Vrije uren", "Totaal uren"])
    style_header(po); po.freeze_panes = "A2"

    summary = fetch_summary_rows()
    for r in summary:
        work = round(float(r["work_hours"] or 0), 2)
        free = round(float(r["free_hours"] or 0), 2)
        total = round(work + free, 2)
        ov.append([r["submitted_at"], r["name"], r["period"], r["week"], work, free, total, "Ja" if r["zero_hours_contract"] else "Nee"])
        wo.append([r["week"], r["period"], r["name"], work, free, total, r["submitted_at"]])
        po.append([r["period"], r["week"], r["name"], work, free, total])

    for ws in (ov, wo, po):
        autosize(ws)

    conn = db()
    employees = conn.execute("SELECT DISTINCT name, name_key FROM submissions ORDER BY name COLLATE NOCASE").fetchall()
    for emp in employees:
        ws = wb.create_sheet(safe_sheet_name(emp["name"]))
        ws.append(["Periode", "Week", "Dag", "Klus / werkadres", "Werkuren", "Vrije uren", "Opmerking", "Ingediend op"])
        style_header(ws); ws.freeze_panes = "A2"
        rows = conn.execute("""
            SELECT s.period, s.week, w.day, w.job, w.work_hours, w.free_hours, w.note, s.submitted_at
            FROM submissions s
            JOIN work_lines w ON w.submission_id=s.id
            WHERE s.name_key=?
            ORDER BY s.week ASC,
                CASE w.day
                    WHEN 'Maandag' THEN 1 WHEN 'Dinsdag' THEN 2 WHEN 'Woensdag' THEN 3
                    WHEN 'Donderdag' THEN 4 WHEN 'Vrijdag' THEN 5 WHEN 'Zaterdag' THEN 6
                    WHEN 'Zondag' THEN 7 ELSE 8 END,
                w.id ASC
        """, (emp["name_key"],)).fetchall()
        for r in rows:
            ws.append([r["period"], r["week"], r["day"], r["job"], r["work_hours"], r["free_hours"], r["note"], r["submitted_at"]])
        autosize(ws)
    conn.close()

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out

@app.on_event("startup")
def startup():
    init_db()

@app.get("/", response_class=HTMLResponse)
def home():
    return (BASE_DIR / "app" / "static" / "index.html").read_text(encoding="utf-8")

@app.get("/admin", response_class=HTMLResponse)
def admin():
    return (BASE_DIR / "app" / "static" / "admin.html").read_text(encoding="utf-8")

@app.post("/api/submit")
def submit_hours(sub: Submission):
    if any(line.day not in DAYS for line in sub.lines):
        raise HTTPException(400, "Ongeldige dag gevonden.")
    result = save_submission(sub)
    return {"ok": True, "message": "Week succesvol ingediend.", **result}

@app.get("/api/status")
def status():
    rows = []
    for r in fetch_summary_rows():
        work = round(float(r["work_hours"] or 0), 2)
        free = round(float(r["free_hours"] or 0), 2)
        rows.append({
            "submitted_at": r["submitted_at"],
            "name": r["name"],
            "period": r["period"],
            "week": r["week"],
            "work_hours": work,
            "free_hours": free,
            "total_hours": round(work + free, 2),
            "zero_hours_contract": "Ja" if r["zero_hours_contract"] else "Nee",
        })
    return rows

@app.get("/api/excel")
def excel_download():
    stream = build_excel()
    headers = {"Content-Disposition": 'attachment; filename="urenregistratie.xlsx"'}
    return StreamingResponse(stream, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers=headers)

@app.get("/health")
def health():
    return {"ok": True, "storage": "sqlite", "excel": "generated-live"}
