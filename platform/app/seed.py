# -*- coding: utf-8 -*-
from .database import Base, engine, SessionLocal
from . import models, kpi_data
from .auth import hash_password

def init_db():
    Base.metadata.create_all(engine)
    db = SessionLocal()
    try:
        if db.query(models.Department).count() > 0:
            return
        dept_by_key = {}
        for name, key, records in kpi_data.DEPARTMENTS:
            d = models.Department(name=name, key=key)
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
                       "realestate":"مدير العقار","digital":"مدير التقنية","hr":"مدير الموارد البشرية"}
        for key, title in role_titles.items():
            mkuser(key, key+"123", "manager", title, key)
        mkuser("emp", "emp123", "employee", "موظف الامتياز", "franchise")
        db.add(models.Setting(key="report_month", value="0"))
        db.add(models.Setting(key="company_name", value="درب · Darb"))
        db.commit()
    finally:
        db.close()
