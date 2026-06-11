# -*- coding: utf-8 -*-
import os, json
from fastapi import FastAPI, Request, Depends, Form, UploadFile, File
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.middleware.sessions import SessionMiddleware
from sqlalchemy.orm import Session

from .database import get_db
from . import models, compute
from .kpi_data import PILLARS, PROJECTS, PROJECT_PILLAR
from .auth import (current_user, verify_password, can_edit_definition, can_edit_value, ROLE_LABEL)
from .seed import init_db

BASE = os.path.dirname(os.path.abspath(__file__))
app = FastAPI(title="منصة درب لمؤشرات الأداء")
app.add_middleware(SessionMiddleware, secret_key=os.environ.get("DARB_SECRET", "darb-secret-key-change-me"))
app.mount("/static", StaticFiles(directory=os.path.join(BASE, "static")), name="static")
templates = Jinja2Templates(directory=os.path.join(BASE, "templates"))
templates.env.globals["fmt_value"] = compute.fmt_value
templates.env.globals["status_class"] = compute.status_class
templates.env.globals["MONTHS"] = compute.MONTHS
templates.env.globals["ROLE_LABEL"] = ROLE_LABEL

init_db()

from types import SimpleNamespace
from .database import SessionLocal
_s = SessionLocal()
templates.env.globals["DEPTS_NAV"] = [SimpleNamespace(id=d.id, name=d.name) for d in _s.query(models.Department).all()]
_s.close()

PERSPECTIVES = ["مالي","العملاء","العمليات الداخلية","التعلّم والنمو"]

def get_month(db):
    s = db.query(models.Setting).get("report_month")
    try: return int(s.value) if s and s.value else 0
    except: return 0

def kpi_calc(kpi, nmonths):
    vals = {v.month: v.actual for v in kpi.values}
    series = [vals.get(m) for m in range(1, 13)]
    y = compute.ytd(kpi.agg, series)
    ach = compute.achievement(kpi.target, kpi.polarity, y)
    fc = compute.forecast(kpi.agg, y, nmonths)
    return {"kpi": kpi, "series": series, "ytd": y, "ach": ach,
            "status": compute.status(ach), "sclass": compute.status_class(ach), "forecast": fc}

def all_calcs(db):
    nmonths = get_month(db)
    kpis = db.query(models.KPI).all()
    return [kpi_calc(k, nmonths) for k in kpis], nmonths

def avg(achs):
    nums = [a for a in achs if a is not None]
    return sum(nums)/len(nums) if nums else None

# ---------------- المصادقة ----------------
@app.get("/", response_class=HTMLResponse)
def root(request: Request):
    return RedirectResponse("/dashboard")

@app.get("/login", response_class=HTMLResponse)
def login_form(request: Request):
    return templates.TemplateResponse(request, "login.html", {"request": request, "error": None})

@app.post("/login")
def login(request: Request, username: str = Form(...), password: str = Form(...), db: Session = Depends(get_db)):
    u = db.query(models.User).filter_by(username=username).first()
    if not u or not verify_password(password, u.pw_hash, u.salt):
        return templates.TemplateResponse(request, "login.html", {"request": request, "error": "بيانات الدخول غير صحيحة"})
    request.session["uid"] = u.id
    return RedirectResponse("/dashboard", status_code=302)

@app.get("/logout")
def logout(request: Request):
    request.session.clear()
    return RedirectResponse("/login", status_code=302)

def require(request, db):
    u = current_user(request, db)
    return u

# ---------------- لوحة الإدارة التنفيذية ----------------
@app.get("/dashboard", response_class=HTMLResponse)
def dashboard(request: Request, db: Session = Depends(get_db)):
    u = require(request, db)
    if not u: return RedirectResponse("/login")
    calcs, nmonths = all_calcs(db)
    achs = [c["ach"] for c in calcs]
    overall = avg(achs)
    counts = {"ok":0,"warn":0,"bad":0,"muted":0}
    for c in calcs: counts[c["sclass"]] += 1
    def rollup(keyfn, keys):
        out = []
        for k in keys:
            sub = [c["ach"] for c in calcs if keyfn(c)==k]
            out.append({"name": k, "ach": avg(sub), "n": len([c for c in calcs if keyfn(c)==k])})
        return out
    pillars = rollup(lambda c: c["kpi"].pillar, PILLARS)
    depts = db.query(models.Department).all()
    dept_roll = []
    for d in depts:
        sub = [c["ach"] for c in calcs if c["kpi"].department_id==d.id]
        dept_roll.append({"id": d.id, "name": d.name, "ach": avg(sub), "n": len(sub)})
    persp = rollup(lambda c: c["kpi"].perspective, PERSPECTIVES)
    projects = []
    for p in PROJECTS:
        sub=[c["ach"] for c in calcs if c["kpi"].project==p]
        projects.append({"name":p,"pillar":PROJECT_PILLAR.get(p,""),"ach":avg(sub),"n":len(sub)})
    chart = {
        "pillars": {"labels": [p["name"] for p in pillars], "data": [round((p["ach"] or 0)*100,1) for p in pillars]},
        "depts": {"labels": [d["name"] for d in dept_roll], "data": [round((d["ach"] or 0)*100,1) for d in dept_roll]},
        "persp": {"labels": [p["name"] for p in persp], "data": [round((p["ach"] or 0)*100,1) for p in persp]},
    }
    return templates.TemplateResponse(request, "dashboard.html", {"request": request, "u": u, "overall": overall,
        "counts": counts, "pillars": pillars, "depts": dept_roll, "persp": persp, "projects": projects,
        "nmonths": nmonths, "chart_json": json.dumps(chart, ensure_ascii=False), "total": len(calcs)})

# ---------------- تغذية مؤشرات إدارة ----------------
@app.get("/department/{dept_id}", response_class=HTMLResponse)
def department(dept_id: int, request: Request, db: Session = Depends(get_db)):
    u = require(request, db)
    if not u: return RedirectResponse("/login")
    d = db.query(models.Department).get(dept_id)
    if not d: return RedirectResponse("/dashboard")
    nmonths = get_month(db)
    rows = []
    for k in d.kpis:
        c = kpi_calc(k, nmonths)
        c["editable"] = can_edit_value(u, k)
        rows.append(c)
    any_edit = any(r["editable"] for r in rows)
    return templates.TemplateResponse(request, "department.html", {"request": request, "u": u, "dept": d,
        "rows": rows, "any_edit": any_edit, "nmonths": nmonths})

@app.post("/department/{dept_id}/save")
async def department_save(dept_id: int, request: Request, db: Session = Depends(get_db)):
    u = require(request, db)
    if not u: return RedirectResponse("/login")
    d = db.query(models.Department).get(dept_id)
    form = await request.form()
    for k in d.kpis:
        if not can_edit_value(u, k):   # حماية: لا تغذية بدون صلاحية
            continue
        for m in range(1, 13):
            field = f"v_{k.id}_{m}"
            if field in form:
                raw = (form.get(field) or "").strip().replace(",", "")
                val = None
                if raw != "":
                    try: val = float(raw)
                    except: val = None
                existing = db.query(models.KPIValue).filter_by(kpi_id=k.id, month=m).first()
                if val is None:
                    if existing: db.delete(existing)
                elif existing:
                    existing.actual = val
                else:
                    db.add(models.KPIValue(kpi_id=k.id, month=m, actual=val))
    db.commit()
    return RedirectResponse(f"/department/{dept_id}?saved=1", status_code=302)

# ---------------- إدارة المؤشرات (أدمن: تعديل المستهدفات والأوزان) ----------------
@app.get("/admin/kpis", response_class=HTMLResponse)
def admin_kpis(request: Request, db: Session = Depends(get_db)):
    u = require(request, db)
    if not u: return RedirectResponse("/login")
    if not can_edit_definition(u):
        return templates.TemplateResponse(request, "forbidden.html", {"request": request, "u": u})
    depts = db.query(models.Department).all()
    return templates.TemplateResponse(request, "admin_kpis.html", {"request": request, "u": u, "depts": depts})

@app.post("/admin/kpis/save")
async def admin_kpis_save(request: Request, db: Session = Depends(get_db)):
    u = require(request, db)
    if not u or not can_edit_definition(u):
        return RedirectResponse("/dashboard")
    form = await request.form()
    for k in db.query(models.KPI).all():
        for attr in ("target", "baseline"):
            f = f"{attr}_{k.id}"
            if f in form:
                raw = (form.get(f) or "").strip().replace(",", "")
                setattr(k, attr, float(raw) if raw != "" else None)
        wf = f"weight_{k.id}"
        if wf in form:
            raw = (form.get(wf) or "").strip().replace(",", "")
            k.weight = (float(raw)/100.0) if raw != "" else 0   # المُدخل نسبة مئوية
        tt = form.get(f"target_text_{k.id}")
        if tt is not None: k.target_text = tt
    db.commit()
    return RedirectResponse("/admin/kpis?saved=1", status_code=302)

# ---------------- الاستراتيجية / الخارطة ----------------
@app.get("/strategy", response_class=HTMLResponse)
def strategy(request: Request, db: Session = Depends(get_db)):
    u = require(request, db)
    if not u: return RedirectResponse("/login")
    calcs, _ = all_calcs(db)
    def roll(keyfn, k):
        sub=[c["ach"] for c in calcs if keyfn(c)==k];
        n=[c for c in calcs if keyfn(c)==k]
        return avg(sub), len(n)
    pillars=[{"name":p, **dict(zip(("ach","n"), roll(lambda c:c["kpi"].pillar,p)))} for p in PILLARS]
    persp=[{"name":p, **dict(zip(("ach","n"), roll(lambda c:c["kpi"].perspective,p)))} for p in PERSPECTIVES]
    projects=[]
    for p in PROJECTS:
        a,n=roll(lambda c:c["kpi"].project,p)
        projects.append({"name":p,"pillar":PROJECT_PILLAR.get(p,""),"ach":a,"n":n,"gap": n==0})
    return templates.TemplateResponse(request, "strategy.html", {"request": request, "u": u,
        "pillars": pillars, "persp": persp, "projects": projects})

# ---------------- المبادرات والمخاطر ----------------
@app.get("/initiatives", response_class=HTMLResponse)
def initiatives(request: Request, db: Session = Depends(get_db)):
    u = require(request, db)
    if not u: return RedirectResponse("/login")
    items = db.query(models.Initiative).all()
    return templates.TemplateResponse(request, "initiatives.html", {"request": request, "u": u, "items": items,
        "pillars": PILLARS, "projects": PROJECTS})

@app.post("/initiatives/save")
async def initiatives_save(request: Request, db: Session = Depends(get_db)):
    u = require(request, db)
    if not u: return RedirectResponse("/login")
    form = await request.form()
    iid = form.get("id")
    def num(x):
        x=(x or "").strip().replace(",","")
        try: return float(x)
        except: return None
    if iid:
        it = db.query(models.Initiative).get(int(iid))
    else:
        it = models.Initiative(name=form.get("name") or "مبادرة")
        db.add(it)
    it.name=form.get("name") or it.name; it.pillar=form.get("pillar") or ""
    it.project=form.get("project") or ""; it.owner=form.get("owner") or ""
    it.start=form.get("start") or ""; it.end=form.get("end") or ""
    it.budget=num(form.get("budget"));
    p=num(form.get("progress")); it.progress=(p/100 if p and p>1 else (p or 0))
    it.notes=form.get("notes") or ""
    db.commit()
    return RedirectResponse("/initiatives", status_code=302)

@app.get("/risks", response_class=HTMLResponse)
def risks(request: Request, db: Session = Depends(get_db)):
    u = require(request, db)
    if not u: return RedirectResponse("/login")
    items = db.query(models.Risk).all()
    return templates.TemplateResponse(request, "risks.html", {"request": request, "u": u, "items": items})

@app.post("/risks/save")
async def risks_save(request: Request, db: Session = Depends(get_db)):
    u = require(request, db)
    if not u: return RedirectResponse("/login")
    form = await request.form()
    rid = form.get("id")
    def int0(x):
        try: return int(x)
        except: return 0
    it = db.query(models.Risk).get(int(rid)) if rid else None
    if not it:
        it = models.Risk(title=form.get("title") or "خطر"); db.add(it)
    it.title=form.get("title") or it.title; it.area=form.get("area") or ""
    it.likelihood=int0(form.get("likelihood")); it.impact=int0(form.get("impact"))
    it.response=form.get("response") or ""; it.owner=form.get("owner") or ""
    db.commit()
    return RedirectResponse("/risks", status_code=302)

# ---------------- إعداد شهر التقرير ----------------
@app.post("/settings/report-month")
async def set_month(request: Request, db: Session = Depends(get_db)):
    u = require(request, db)
    if not u or u.role not in ("admin","executive"):
        return RedirectResponse("/dashboard")
    form = await request.form()
    s = db.query(models.Setting).get("report_month")
    if not s: s = models.Setting(key="report_month"); db.add(s)
    s.value = (form.get("report_month") or "0")
    db.commit()
    return RedirectResponse("/dashboard", status_code=302)

# ---------------- رفع/استيراد من إكسل (أدمن) ----------------
@app.get("/upload", response_class=HTMLResponse)
def upload_form(request: Request, db: Session = Depends(get_db)):
    u = require(request, db)
    if not u: return RedirectResponse("/login")
    if not can_edit_definition(u):
        return templates.TemplateResponse(request, "forbidden.html", {"request": request, "u": u})
    return templates.TemplateResponse(request, "upload.html", {"request": request, "u": u, "msg": None})

@app.post("/upload", response_class=HTMLResponse)
async def upload_import(request: Request, file: UploadFile = File(...), db: Session = Depends(get_db)):
    u = require(request, db)
    if not u or not can_edit_definition(u):
        return RedirectResponse("/dashboard")
    import openpyxl, io
    data = await file.read()
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    name_to_kpi = {k.name: k for k in db.query(models.KPI).all()}
    imported = 0
    for ws in wb.worksheets:
        if "المؤشرات" not in ws.title:  # أوراق المؤشرات فقط
            continue
        for row in ws.iter_rows(min_row=4, values_only=True):
            if len(row) < 20 or not row[2]:  # العمود C = اسم المؤشر
                continue
            k = name_to_kpi.get(str(row[2]).strip())
            if not k: continue
            for m in range(12):
                cell = row[8+m]  # I..T = الأشهر
                if isinstance(cell, (int, float)):
                    ex = db.query(models.KPIValue).filter_by(kpi_id=k.id, month=m+1).first()
                    if ex: ex.actual = float(cell)
                    else: db.add(models.KPIValue(kpi_id=k.id, month=m+1, actual=float(cell)))
                    imported += 1
    db.commit()
    return templates.TemplateResponse(request, "upload.html", {"request": request, "u": u,
        "msg": f"تم استيراد {imported} قيمة شهرية من الملف."})
