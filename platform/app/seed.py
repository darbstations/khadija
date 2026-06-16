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
            seed_eval_forms(db, dbk); seed_marketing_staff(db, dbk); db.commit()
            return
        dept_by_key = {}
        dept_w = round(1.0/len(kpi_data.DEPARTMENTS), 4)
        for name, key, records in kpi_data.DEPARTMENTS:
            d = models.Department(name=name, key=key, weight=dept_w)
            db.add(d); db.flush()
            dept_by_key[key] = d
            n = len(records); w = round(1.0/n, 4) if n else 0
            for i, rec in enumerate(records):
                axis, nm, unit, pol, agg, tgt, ttxt, fmt, pillar, project = rec
                db.add(models.KPI(
                    department_id=d.id, order=i, axis=axis, name=nm, unit=unit,
                    polarity=pol, agg=agg, target=tgt, target_text=ttxt, fmt=fmt,
                    pillar=pillar, project=project, perspective=kpi_data.perspective(nm),
                    kpitype=kpi_data.classify(nm), weight=w))
        # المستخدمون (صلاحيات)
        def mkuser(username, pw, role, full, dept_key=None):
            h, s = hash_password(pw)
            db.add(models.User(username=username, pw_hash=h, salt=s, role=role,
                               full_name=full, department_id=dept_by_key[dept_key].id if dept_key else None))
        mkuser("admin", "admin123", "admin", "مدير النظام")
        mkuser("exec", "exec123", "executive", "الرئيس التنفيذي", "exec")
        role_titles = {"franchise":"مدير الامتياز","operations":"مدير التشغيل","investment":"مدير الاستثمار",
                       "realestate":"مدير العقار","digital":"مدير التقنية","hr":"مدير الموارد البشرية",
                       "marketing":"مدير التسويق"}
        for key, title in role_titles.items():
            mkuser(key, key+"123", "manager", title, key)
        mkuser("emp", "emp123", "employee", "موظف الامتياز", "franchise")
        db.add(models.Setting(key="report_month", value="0"))
        db.add(models.Setting(key="company_name", value="درب · Darb"))
        seed_eval_forms(db, dept_by_key)
        seed_marketing_staff(db, dept_by_key)
        db.commit()
    finally:
        db.close()
