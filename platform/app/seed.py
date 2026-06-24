# -*- coding: utf-8 -*-
from .database import Base, engine, SessionLocal
from . import models, kpi_data, eval_data, marketing_staff
from .auth import hash_password

def seed_marketing_staff(db, dept_by_key):
    if db.query(models.User).filter_by(username="amani").first():
        return
    mk = dept_by_key.get("marketing")
    if not mk: return
    for username, full, role, section, items in marketing_staff.STAFF:
        h, s = hash_password(username + "123")
        u = models.User(username=username, pw_hash=h, salt=s, role="employee",
                        full_name=full, department_id=mk.id)
        db.add(u); db.flush()
        ew = round(1.0/len(items), 4) if items else 0
        for nm, fmt, tgt, ach in items:
            pol = "↑"
            k = models.KPI(department_id=mk.id, order=0, axis=section, name=f"{full} — {nm}",
                           unit="", polarity=pol, agg="LAST", target=tgt,
                           target_text=("100%" if (fmt=="pct" and tgt==1.0) else (str(tgt) if tgt is not None else "—")),
                           fmt=fmt, pillar="", project="", perspective="", kpitype="",
                           weight=ew, level="individual", section=section, owner_user_id=u.id)
            db.add(k); db.flush()
            if ach is not None:
                db.add(models.KPIValue(kpi_id=k.id, month=6, actual=float(ach)))

def seed_initiatives(db):
    if db.query(models.Initiative).count() > 0:
        return
    for nm, qty, qtype, done, pct in getattr(kpi_data, "IT_PROJECTS", []):
        db.add(models.Initiative(name=nm, pillar=kpi_data.PL_INNO, project=kpi_data.P_TECH,
            kpi_link="نسبة إنجاز محفظة مشاريع التحول الرقمي", owner="إدارة التقنية",
            progress=float(pct), notes=f"{done} من {qty} {qtype}"))

def seed_eval_forms(db, dept_by_key):
    if db.query(models.EvalForm).count() > 0:
        return
    for dept_key, code, role, items in eval_data.EVAL_FORMS:
        d = dept_by_key.get(dept_key)
        if not d: continue
        f = models.EvalForm(department_id=d.id, code=code, role=role)
        db.add(f); db.flush()
        for i, (nm, w, tt) in enumerate(items):
            db.add(models.EvalItem(form_id=f.id, order=i, name=nm, weight=w/100.0, target_text=tt))

def init_db():
    Base.metadata.create_all(engine)
    db = SessionLocal()
    try:
        if db.query(models.Department).count() > 0:
            dbk = {d.key: d for d in db.query(models.Department).all()}
            seed_eval_forms(db, dbk); seed_marketing_staff(db, dbk); seed_initiatives(db); db.commit()
            return
        dept_by_key = {}
        for name, key, records in kpi_data.DEPARTMENTS:
            d = models.Department(name=name, key=key, weight=kpi_data.DEPT_WEIGHTS.get(key, 0))
            db.add(d); db.flush()
            dept_by_key[key] = d
            ws = kpi_data.kpi_weights(records)
            for i, rec in enumerate(records):
                axis, nm, unit, pol, agg, tgt, ttxt, fmt, pillar, project = rec
                k = models.KPI(
                    department_id=d.id, order=i, axis=axis, name=nm, unit=unit,
                    polarity=pol, agg=agg, target=tgt, target_text=ttxt, fmt=fmt,
                    pillar=pillar, project=project, perspective=kpi_data.perspective(nm),
                    kpitype=kpi_data.priority_label(kpi_data.priority(nm)), weight=ws[i])
                db.add(k); db.flush()
                if nm in kpi_data.SEED_ACTUALS:   # القيم المُحقَّقة من الملفات
                    db.add(models.KPIValue(kpi_id=k.id, month=6, actual=float(kpi_data.SEED_ACTUALS[nm])))
        # المستخدمون (صلاحيات)
        def mkuser(username, pw, role, full, dept_key=None):
            h, s = hash_password(pw)
            db.add(models.User(username=username, pw_hash=h, salt=s, role=role,
                               full_name=full, department_id=dept_by_key[dept_key].id if dept_key else None))
        mkuser("admin", "admin123", "admin", "مدير النظام")
        mkuser("exec", "exec123", "executive", "الرئيس التنفيذي", "exec")
        role_titles = {"franchise":"مدير الامتياز","operations":"مدير التشغيل","investment":"مدير الاستثمار",
                       "realestate":"مدير العقار","digital":"مدير التقنية","hr":"مدير الموارد البشرية",
                       "marketing":"مدير التسويق","quality":"مدير الجودة","legal":"المدير القانوني"}
        for key, title in role_titles.items():
            mkuser(key, key+"123", "manager", title, key)
        mkuser("emp", "emp123", "employee", "موظف الامتياز", "franchise")
        db.add(models.Setting(key="report_month", value="0"))
        db.add(models.Setting(key="company_name", value="درب · Darb"))
        seed_eval_forms(db, dept_by_key)
        seed_marketing_staff(db, dept_by_key)
        seed_initiatives(db)
        db.commit()
    finally:
        db.close()
