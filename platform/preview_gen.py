# -*- coding: utf-8 -*-
"""يولّد معاينة كاملة مترابطة لمنصة درب (CSS وشعار مضمّنان + روابط محلية) مع بيانات تجريبية.
   التشغيل:  cd platform && python3 preview_gen.py   ·  المخرجات: ../preview/  (افتح index.html)"""
import os, base64, random, re

os.environ.setdefault("DARB_DB_PATH", "/tmp/darb_preview.db")
os.environ.setdefault("DARB_SECRET", "draft")
if os.path.exists(os.environ["DARB_DB_PATH"]):
    os.remove(os.environ["DARB_DB_PATH"])

from app.seed import init_db
init_db()
from app.database import SessionLocal
from app import models

db = SessionLocal()
db.query(models.Setting).filter_by(key="report_month").update({"value": "6"})
random.seed(11)
# تعبئة المؤشرات الاستراتيجية بأرقام تجريبية (أشهر 1..6)
for k in db.query(models.KPI).filter_by(level="strategic").all():
    base = k.target
    for m in range(1, 7):
        if base is None: val = random.choice([60, 120, 500, 2000])
        elif k.agg == "SUM": val = round(base/12*random.uniform(0.6, 1.05), 2)
        else: val = round(base*random.uniform(0.65, 1.03), 4)
        db.add(models.KPIValue(kpi_id=k.id, month=m, actual=val))
db.commit()
# القيم الفعلية لمؤشرات التسويق الاستراتيجية من التقرير (تغلب على العشوائي)
mk = db.query(models.Department).filter_by(key="marketing").first()
ACH = {"الوصول الكلي (Total Reach)":5000000,"عملاء تجاريون جدد":1,"إجمالي المتابعين (كل المنصات)":253000,
 "معدل التفاعل (Engagement)":0,"حملات تسويقية منجزة":11,"تحويل الفرص إلى عقود":0.17,
 "نسبة المبيعات لكل حملة":0.16,"نسبة التزام الهوية البصرية":0.10,"وصول العلاقات العامة (PR)":1500000,
 "عدد المعارض بالسنة":0,"اتفاقيات استراتيجية":0,"تقييم خرائط قوقل":4.0,"معدل خفض الشكاوى":0.21}
mk_ids = [k.id for k in mk.kpis if k.level=="strategic" and k.name in ACH]
db.query(models.KPIValue).filter(models.KPIValue.kpi_id.in_(mk_ids)).delete(synchronize_session=False)
db.flush()
for k in mk.kpis:
    if k.id in mk_ids:
        db.add(models.KPIValue(kpi_id=k.id, month=6, actual=float(ACH[k.name])))
db.commit()

from fastapi.testclient import TestClient
from app.main import app
c = TestClient(app); c.post("/login", data={"username": "admin", "password": "admin123"})

HERE = os.path.dirname(os.path.abspath(__file__))
css = open(os.path.join(HERE, "app/static/style.css"), encoding="utf-8").read()
logo = "data:image/png;base64," + base64.b64encode(open(os.path.join(HERE, "app/static/darb_logo.png"), "rb").read()).decode()

depts = db.query(models.Department).all()
emps = db.query(models.User).filter_by(department_id=mk.id, role="employee").all()

def localize(html):
    html = html.replace('<link rel="stylesheet" href="/static/style.css">', "<style>\n"+css+"\n</style>")
    html = html.replace("/static/darb_logo.png", logo)
    repl = {'href="/dashboard"':'href="dashboard.html"','href="/strategy"':'href="strategy.html"',
        'href="/initiatives"':'href="initiatives.html"','href="/risks"':'href="risks.html"',
        'href="/employees"':'href="employees.html"','href="/evaluations"':'href="evaluations.html"',
        'href="/report/print"':'href="report.html"','href="/admin/kpis"':'href="admin_kpis.html"',
        'href="/admin/staff"':'href="admin_staff.html"','href="/upload"':'href="upload.html"','href="/logout"':'href="index.html"'}
    for a,b in repl.items(): html = html.replace(a,b)
    html = re.sub(r'href="/department/(\d+)"', r'href="dept_\1.html"', html)
    html = re.sub(r'href="/employee/(\d+)"', r'href="emp_\1.html"', html)
    return html

out = os.path.join(HERE, "..", "preview"); os.makedirs(out, exist_ok=True)
def save(name, url):
    r = c.get(url); open(os.path.join(out, name), "w", encoding="utf-8").write(localize(r.text)); print(name, r.status_code)

pages = {"dashboard.html":"/dashboard","strategy.html":"/strategy","employees.html":"/employees",
 "evaluations.html":"/evaluations","initiatives.html":"/initiatives","risks.html":"/risks",
 "admin_kpis.html":"/admin/kpis","admin_staff.html":"/admin/staff","report.html":"/report/print","upload.html":"/upload"}
for n,u in pages.items(): save(n,u)
for d in depts: save(f"dept_{d.id}.html", f"/department/{d.id}")
for e in emps: save(f"emp_{e.id}.html", f"/employee/{e.id}")

# صفحة رئيسية (خريطة المنصة)
cards = [("📊 اللوحة التنفيذية","dashboard.html","الإنجاز العام + الركائز + الإدارات + الرسوم"),
 ("🧭 الاستراتيجية والخارطة","strategy.html","الركائز ومنظورات BSC وتغطية المشاريع"),
 ("👥 الموظفون","employees.html","إنجاز كل موظف وتصنيفه"),
 ("📝 تقييم الموظفين","evaluations.html","نماذج تقييم موزونة"),
 ("🚀 المبادرات","initiatives.html","ربط النتائج بالتنفيذ"),
 ("⚠️ المخاطر","risks.html","سجل المخاطر"),
 ("🛠️ إدارة المؤشرات","admin_kpis.html","تعديل المستهدفات والأوزان (أدمن)"),
 ("👤 إدارة الموظفين","admin_staff.html","إضافة موظفين لأي إدارة (أدمن)"),
 ("🖨️ تقرير للطباعة","report.html","تقرير PDF"),
 ("⬆️ استيراد إكسل","upload.html","تغذية دفعة واحدة")]
dept_links = "".join(f'<a class="d" href="dept_{d.id}.html">{d.name}</a>' for d in depts)
card_html = "".join(f'<a class="card" href="{u}"><div class="t">{t}</div><div class="s">{s}</div></a>' for t,u,s in cards)
index = f"""<!DOCTYPE html><html lang="ar" dir="rtl"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>منصة درب — خريطة</title>
<style>{css}
body{{padding:0}} .hero{{background:var(--grey-d);color:#fff;padding:26px;display:flex;align-items:center;gap:18px}}
.hero img{{height:54px;background:#fff;border-radius:10px;padding:5px 10px}}
.wrap{{max-width:1100px;margin:24px auto;padding:0 18px}}
.cards2{{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:14px}}
a.card{{background:#fff;border:1px solid var(--line);border-radius:14px;padding:18px;display:block;border-top:4px solid var(--orange)}}
a.card:hover{{box-shadow:0 4px 14px rgba(0,0,0,.08)}}
a.card .t{{font-weight:800;color:var(--grey-d);font-size:16px}} a.card .s{{color:var(--grey);font-size:13px;margin-top:6px}}
.depts{{margin-top:18px;display:flex;gap:8px;flex-wrap:wrap}} a.d{{background:var(--grey);color:#fff;padding:8px 12px;border-radius:8px;font-size:13px}}
a.d:hover{{background:var(--orange)}}
</style></head><body>
<div class="hero"><img src="{logo}"><div><h1 style="margin:0;color:#fff">منصة درب لمؤشرات الأداء — خريطة المنصة</h1>
<div style="color:#ccc;font-size:13px">معاينة مترابطة · افتح أي بطاقة للتنقّل · بيانات تجريبية</div></div></div>
<div class="wrap">
<div class="cards2">{card_html}</div>
<h2>تغذية مؤشرات الإدارات</h2><div class="depts">{dept_links}</div>
<h2>موظفو التسويق (مؤشرات فردية)</h2><div class="depts">{''.join(f'<a class="d" href="emp_{e.id}.html">{e.full_name}</a>' for e in emps)}</div>
<p class="small" style="margin-top:22px">احفظ كل الملفات بنفس المجلد لتعمل الروابط · درب · Darb</p>
</div></body></html>"""
open(os.path.join(out, "index.html"), "w", encoding="utf-8").write(index)
print("index.html ✅  →", os.path.abspath(out))
