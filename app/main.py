from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field
from typing import List, Optional
from pathlib import Path
from datetime import datetime
import re
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)
WORKBOOK = DATA_DIR / "urenregistratie.xlsx"

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
        ws.column_dimensions[get_column_letter(col)].width = min(max(width + 2, 10), 34)

def ensure_summary_sheets(wb):
    if "Overzicht" not in wb.sheetnames:
        ws = wb.create_sheet("Overzicht", 0)
        ws.append(["Ingediend op", "Naam", "Periode", "Week", "Werkuren", "Vrije uren", "Totaal uren", "0-urencontract"])
        style_header(ws); ws.freeze_panes = "A2"
    if "Week-overzicht" not in wb.sheetnames:
        ws = wb.create_sheet("Week-overzicht", 1)
        ws.append(["Week", "Periode", "Naam", "Werkuren", "Vrije uren", "Totaal uren", "Ingediend op"])
        style_header(ws); ws.freeze_panes = "A2"
    if "Periode-overzicht" not in wb.sheetnames:
        ws = wb.create_sheet("Periode-overzicht", 2)
        ws.append(["Periode", "Week", "Naam", "Werkuren", "Vrije uren", "Totaal uren"])
        style_header(ws); ws.freeze_panes = "A2"

def init_workbook():
    if WORKBOOK.exists():
        wb = load_workbook(WORKBOOK)
        ensure_summary_sheets(wb)
        wb.save(WORKBOOK)
        return
    wb = Workbook()
    ws = wb.active
    ws.title = "Overzicht"
    ws.append(["Ingediend op", "Naam", "Periode", "Week", "Werkuren", "Vrije uren", "Totaal uren", "0-urencontract"])
    style_header(ws); ws.freeze_panes = "A2"
    wk = wb.create_sheet("Week-overzicht")
    wk.append(["Week", "Periode", "Naam", "Werkuren", "Vrije uren", "Totaal uren", "Ingediend op"])
    style_header(wk); wk.freeze_panes = "A2"
    ps = wb.create_sheet("Periode-overzicht")
    ps.append(["Periode", "Week", "Naam", "Werkuren", "Vrije uren", "Totaal uren"])
    style_header(ps); ps.freeze_panes = "A2"
    wb.save(WORKBOOK)

def upsert_submission(sub: Submission):
    init_workbook(); wb = load_workbook(WORKBOOK); ensure_summary_sheets(wb)
    period = period_for_week(sub.week); name = sub.name.strip(); sheet_name = safe_sheet_name(name)
    if sheet_name not in wb.sheetnames:
        ws = wb.create_sheet(sheet_name)
        ws.append(["Periode", "Week", "Dag", "Klus / werkadres", "Werkuren", "Vrije uren", "Opmerking", "Ingediend op"])
        style_header(ws); ws.freeze_panes = "A2"
    ws = wb[sheet_name]
    for r in reversed([r for r in range(2, ws.max_row + 1) if ws.cell(r, 2).value == sub.week]): ws.delete_rows(r)
    ts = datetime.now().strftime("%d-%m-%Y %H:%M"); work_total = 0.0; free_total = 0.0
    for line in sub.lines:
        if line.hours == 0 and line.free_hours == 0: continue
        work = float(line.hours); free = float(line.free_hours) if sub.zero_hours_contract else 0.0
        work_total += work; free_total += free
        ws.append([period, sub.week, line.day, line.job, work, free, line.note or "", ts])
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 5).number_format = "0.00"; ws.cell(r, 6).number_format = "0.00"
    autosize(ws)

    ov = wb["Overzicht"]
    for r in range(ov.max_row, 1, -1):
        if str(ov.cell(r, 2).value).strip().lower() == name.lower() and ov.cell(r, 4).value == sub.week: ov.delete_rows(r)
    ov.append([ts, name, period, sub.week, round(work_total, 2), round(free_total, 2), round(work_total + free_total, 2), "Ja" if sub.zero_hours_contract else "Nee"])
    autosize(ov)

    wo = wb["Week-overzicht"]
    for r in range(wo.max_row, 1, -1):
        if str(wo.cell(r, 3).value).strip().lower() == name.lower() and wo.cell(r, 1).value == sub.week: wo.delete_rows(r)
    wo.append([sub.week, period, name, round(work_total, 2), round(free_total, 2), round(work_total + free_total, 2), ts])
    rows = list(wo.iter_rows(min_row=2, values_only=True))
    if wo.max_row > 1: wo.delete_rows(2, wo.max_row - 1)
    for row in sorted(rows, key=lambda x: (x[0] or 0, str(x[2] or ""))): wo.append(row)
    autosize(wo)

    po = wb["Periode-overzicht"]
    for r in range(po.max_row, 1, -1):
        if str(po.cell(r, 3).value).strip().lower() == name.lower() and po.cell(r, 2).value == sub.week: po.delete_rows(r)
    po.append([period, sub.week, name, round(work_total, 2), round(free_total, 2), round(work_total + free_total, 2)])
    autosize(po)

    wb.save(WORKBOOK)
    return {"period": period, "work_total": round(work_total, 2), "free_total": round(free_total, 2)}

@app.get("/", response_class=HTMLResponse)
def home(): return (BASE_DIR / "app" / "static" / "index.html").read_text(encoding="utf-8")

@app.get("/admin", response_class=HTMLResponse)
def admin(): return (BASE_DIR / "app" / "static" / "admin.html").read_text(encoding="utf-8")

@app.post("/api/submit")
def submit_hours(sub: Submission):
    if not sub.lines: raise HTTPException(400, "Vul minimaal 1 urenregel in.")
    if any(line.day not in DAYS for line in sub.lines): raise HTTPException(400, "Ongeldige dag gevonden.")
    result = upsert_submission(sub)
    return {"ok": True, "message": "Week succesvol ingediend.", **result}

@app.get("/api/status")
def status():
    init_workbook(); wb = load_workbook(WORKBOOK, data_only=True); ov = wb["Overzicht"]; rows = []
    for r in range(2, ov.max_row + 1):
        if ov.cell(r, 2).value:
            rows.append({"submitted_at": ov.cell(r,1).value,"name": ov.cell(r,2).value,"period": ov.cell(r,3).value,"week": ov.cell(r,4).value,"work_hours": ov.cell(r,5).value,"free_hours": ov.cell(r,6).value,"total_hours": ov.cell(r,7).value,"zero_hours_contract": ov.cell(r,8).value})
    rows.sort(key=lambda x: (x["week"], x["name"]), reverse=True); return rows

@app.get("/api/excel")
def excel_download():
    init_workbook(); return FileResponse(WORKBOOK, filename="urenregistratie.xlsx", media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.get("/health")
def health(): return {"ok": True}
