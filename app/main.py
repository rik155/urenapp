from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field
from typing import List, Optional
from pathlib import Path
from datetime import datetime
from io import BytesIO
import re, sqlite3
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE_DIR=Path(__file__).resolve().parent.parent; DATA_DIR=BASE_DIR/'data'; DATA_DIR.mkdir(exist_ok=True); DB_PATH=DATA_DIR/'urenregistratie.db'; LEGACY_WORKBOOK=DATA_DIR/'urenregistratie.xlsx'
app=FastAPI(title='UrenApp'); app.mount('/static',StaticFiles(directory=str(BASE_DIR/'app'/'static')),name='static')
DAYS=['Maandag','Dinsdag','Woensdag','Donderdag','Vrijdag','Zaterdag','Zondag']
class WorkLine(BaseModel):
    day:str; job:str=Field(min_length=1,max_length=120); hours:float=Field(ge=0,le=24); free_hours:float=Field(default=0,ge=0,le=24); note:Optional[str]=Field(default='',max_length=250)
class Submission(BaseModel):
    name:str=Field(min_length=2,max_length=60); week:int=Field(ge=1,le=53); zero_hours_contract:bool=False; lines:List[WorkLine]
def period_for_week(w): return ((w-1)//4)+1
def safe_sheet_name(n): return (re.sub(r'[\\/*?:\[\]]','-',n).strip() or 'Medewerker')[:31]
def db():
    c=sqlite3.connect(DB_PATH,timeout=30); c.row_factory=sqlite3.Row; c.execute('PRAGMA foreign_keys=ON'); return c
def init_db():
    c=db(); c.executescript('''PRAGMA journal_mode=WAL; CREATE TABLE IF NOT EXISTS submissions(id INTEGER PRIMARY KEY AUTOINCREMENT,name TEXT NOT NULL,name_key TEXT NOT NULL,week INTEGER NOT NULL,period INTEGER NOT NULL,zero_hours_contract INTEGER NOT NULL DEFAULT 0,submitted_at TEXT NOT NULL,UNIQUE(name_key,week)); CREATE TABLE IF NOT EXISTS work_lines(id INTEGER PRIMARY KEY AUTOINCREMENT,submission_id INTEGER NOT NULL,day TEXT NOT NULL,job TEXT NOT NULL,work_hours REAL NOT NULL DEFAULT 0,free_hours REAL NOT NULL DEFAULT 0,note TEXT NOT NULL DEFAULT '',FOREIGN KEY(submission_id) REFERENCES submissions(id) ON DELETE CASCADE);'''); c.commit(); c.close(); migrate_legacy_excel_if_needed()
def migrate_legacy_excel_if_needed():
    if not LEGACY_WORKBOOK.exists(): return
    c=db()
    if c.execute('SELECT COUNT(*) FROM submissions').fetchone()[0]>0: c.close(); return
    try:
        wb=load_workbook(LEGACY_WORKBOOK,data_only=True)
        for sn in wb.sheetnames:
            if sn in {'Overzicht','Week-overzicht','Periode-overzicht'}: continue
            ws=wb[sn]; grouped={}
            for r in range(2,ws.max_row+1):
                week,day,job=ws.cell(r,2).value,ws.cell(r,3).value,ws.cell(r,4).value
                if not week or not day or not job: continue
                grouped.setdefault(int(week),[]).append((str(day),str(job),float(ws.cell(r,5).value or 0),float(ws.cell(r,6).value or 0),str(ws.cell(r,7).value or ''),str(ws.cell(r,8).value or datetime.now().strftime('%d-%m-%Y %H:%M'))))
            for week,lines in grouped.items():
                cur=c.execute('INSERT OR IGNORE INTO submissions(name,name_key,week,period,zero_hours_contract,submitted_at) VALUES(?,?,?,?,?,?)',(sn,sn.strip().lower(),week,period_for_week(week),1 if any(x[3]>0 for x in lines) else 0,lines[-1][5]))
                if cur.lastrowid:
                    for x in lines:c.execute('INSERT INTO work_lines(submission_id,day,job,work_hours,free_hours,note) VALUES(?,?,?,?,?,?)',(cur.lastrowid,x[0],x[1],x[2],x[3],x[4]))
        c.commit()
    finally:c.close()
def save_submission(s):
    name=s.name.strip(); key=name.lower(); valid=[x for x in s.lines if x.hours!=0 or x.free_hours!=0]
    if not valid: raise HTTPException(400,'Vul minimaal 1 urenregel in.')
    c=db()
    try:
        c.execute('BEGIN IMMEDIATE'); old=c.execute('SELECT id FROM submissions WHERE name_key=? AND week=?',(key,s.week)).fetchone()
        if old:c.execute('DELETE FROM submissions WHERE id=?',(old['id'],))
        stamp=datetime.now().strftime('%d-%m-%Y %H:%M'); cur=c.execute('INSERT INTO submissions(name,name_key,week,period,zero_hours_contract,submitted_at) VALUES(?,?,?,?,?,?)',(name,key,s.week,period_for_week(s.week),1 if s.zero_hours_contract else 0,stamp)); wt=ft=0.0
        for x in valid:
            w=float(x.hours); f=float(x.free_hours) if s.zero_hours_contract else 0.0; wt+=w; ft+=f; c.execute('INSERT INTO work_lines(submission_id,day,job,work_hours,free_hours,note) VALUES(?,?,?,?,?,?)',(cur.lastrowid,x.day,x.job.strip(),w,f,(x.note or '').strip()))
        c.commit(); return {'period':period_for_week(s.week),'work_total':round(wt,2),'free_total':round(ft,2)}
    except Exception:c.rollback(); raise
    finally:c.close()
def style_header(ws):
    fill=PatternFill('solid',fgColor='1F4E78')
    for x in ws[1]:x.font=Font(color='FFFFFF',bold=True);x.fill=fill;x.alignment=Alignment(vertical='center')
    ws.row_dimensions[1].height=22
def autosize(ws):
    d={}
    for row in ws.iter_rows():
        for x in row:
            if x.value is not None:d[x.column]=max(d.get(x.column,0),len(str(x.value)))
    for col,w in d.items():ws.column_dimensions[get_column_letter(col)].width=min(max(w+2,10),36)
def fetch_summary_rows():
    c=db(); r=c.execute('''SELECT s.id,s.submitted_at,s.name,s.name_key,s.week,s.period,s.zero_hours_contract,COALESCE(SUM(w.work_hours),0) work_hours,COALESCE(SUM(w.free_hours),0) free_hours FROM submissions s LEFT JOIN work_lines w ON w.submission_id=s.id GROUP BY s.id ORDER BY s.week DESC,s.name COLLATE NOCASE''').fetchall();c.close();return r
def build_excel():
    wb=Workbook();ov=wb.active;ov.title='Overzicht';ov.append(['Ingediend op','Naam','Periode','Week','Werkuren','Vrije uren','Totaal uren','0-urencontract']);style_header(ov);ov.freeze_panes='A2'
    wo=wb.create_sheet('Week-overzicht');wo.append(['Week','Periode','Naam','Werkuren','Vrije uren','Totaal uren','Ingediend op']);style_header(wo);wo.freeze_panes='A2'
    po=wb.create_sheet('Periode-overzicht');po.append(['Periode','Week','Naam','Werkuren','Vrije uren','Totaal uren']);style_header(po);po.freeze_panes='A2'
    summary=fetch_summary_rows()
    for r in summary:
        w=round(float(r['work_hours']),2);f=round(float(r['free_hours']),2);t=round(w+f,2);ov.append([r['submitted_at'],r['name'],r['period'],r['week'],w,f,t,'Ja' if r['zero_hours_contract'] else 'Nee']);wo.append([r['week'],r['period'],r['name'],w,f,t,r['submitted_at']]);po.append([r['period'],r['week'],r['name'],w,f,t])
    for ws in (ov,wo,po):autosize(ws)
    c=db()
    for emp in c.execute('SELECT DISTINCT name,name_key FROM submissions ORDER BY name COLLATE NOCASE').fetchall():
        ws=wb.create_sheet(safe_sheet_name(emp['name']));ws.append(['Periode','Week','Dag','Klus / werkadres','Werkuren','Vrije uren','Opmerking','Ingediend op']);style_header(ws);ws.freeze_panes='A2'
        rr=c.execute('''SELECT s.period,s.week,w.day,w.job,w.work_hours,w.free_hours,w.note,s.submitted_at FROM submissions s JOIN work_lines w ON w.submission_id=s.id WHERE s.name_key=? ORDER BY s.week,CASE w.day WHEN 'Maandag' THEN 1 WHEN 'Dinsdag' THEN 2 WHEN 'Woensdag' THEN 3 WHEN 'Donderdag' THEN 4 WHEN 'Vrijdag' THEN 5 WHEN 'Zaterdag' THEN 6 WHEN 'Zondag' THEN 7 ELSE 8 END,w.id''',(emp['name_key'],)).fetchall()
        for r in rr:ws.append([r['period'],r['week'],r['day'],r['job'],r['work_hours'],r['free_hours'],r['note'],r['submitted_at']])
        autosize(ws)
    c.close();out=BytesIO();wb.save(out);out.seek(0);return out
@app.on_event('startup')
def startup():init_db()
@app.get('/',response_class=HTMLResponse)
def home():return (BASE_DIR/'app'/'static'/'index.html').read_text(encoding='utf-8')
@app.get('/admin',response_class=HTMLResponse)
def admin():return (BASE_DIR/'app'/'static'/'admin.html').read_text(encoding='utf-8')
@app.post('/api/submit')
def submit_hours(s:Submission):
    if any(x.day not in DAYS for x in s.lines):raise HTTPException(400,'Ongeldige dag gevonden.')
    return {'ok':True,'message':'Week succesvol ingediend.',**save_submission(s)}
@app.get('/api/status')
def status():
    out=[]
    for r in fetch_summary_rows():
        w=round(float(r['work_hours']),2);f=round(float(r['free_hours']),2);out.append({'id':r['id'],'submitted_at':r['submitted_at'],'name':r['name'],'name_key':r['name_key'],'period':r['period'],'week':r['week'],'work_hours':w,'free_hours':f,'total_hours':round(w+f,2),'zero_hours_contract':'Ja' if r['zero_hours_contract'] else 'Nee'})
    return out
@app.delete('/api/submission/{sid}')
def delete_submission(sid:int):
    c=db();r=c.execute('SELECT name,week FROM submissions WHERE id=?',(sid,)).fetchone()
    if not r:c.close();raise HTTPException(404,'Inzending niet gevonden.')
    c.execute('DELETE FROM submissions WHERE id=?',(sid,));c.commit();c.close();return {'ok':True,'message':f"Week {r['week']} van {r['name']} is verwijderd."}
@app.delete('/api/employee')
def delete_employee(name_key:str=Query(...)):
    c=db();r=c.execute('SELECT name FROM submissions WHERE name_key=? LIMIT 1',(name_key,)).fetchone()
    if not r:c.close();raise HTTPException(404,'Medewerker niet gevonden.')
    name=r['name'];c.execute('DELETE FROM submissions WHERE name_key=?',(name_key,));c.commit();c.close();return {'ok':True,'message':f'{name} en alle weken zijn verwijderd.'}
@app.get('/api/excel')
def excel_download():return StreamingResponse(build_excel(),media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={'Content-Disposition':'attachment; filename="urenregistratie.xlsx"'})
@app.get('/health')
def health():return {'ok':True,'storage':'sqlite','excel':'generated-live'}
