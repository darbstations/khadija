# -*- coding: utf-8 -*-
"""يولّد معاينة HTML مستقلة (CSS وشعار مضمّنان) لصفحات المنصة مع بيانات تجريبية.
   التشغيل من داخل مجلد platform:  python3 preview_gen.py
   المخرجات في ../preview/*.html — افتحها بأي متصفح."""
import os, base64, random

os.environ.setdefault("DARB_DB_PATH", "/tmp/darb_demo.db")
os.environ.setdefault("DARB_SECRET", "draft")
if os.path.exists(os.environ["DARB_DB_PATH"]):
    os.remove(os.environ["DARB_DB_PATH"])

from app.seed import init_db
init_db()
from app.database import SessionLocal
from app import models

db = SessionLocal()
db.query(models.Setting).filter_by(key="report_month").update({"value": "4"})
random.seed(7)
for k in db.query(models.KPI).all():
    base = k.target
    for m in range(1, 5):  # يناير–أبريل بيانات تجريبية
        if base is None:
            val = random.choice([60, 90, 150, 500, 1200])
        elif k.agg == "SUM":
            val = round(base / 12 * random.uniform(0.72, 1.05), 2)
        else:
            val = round(base * random.uniform(0.82, 1.04), 4)
        db.add(models.KPIValue(kpi_id=k.id, month=m, actual=val))
db.commit()

from fastapi.testclient import TestClient
from app.main import app
c = TestClient(app)
c.post("/login", data={"username": "admin", "password": "admin123"})

HERE = os.path.dirname(os.path.abspath(__file__))
css = open(os.path.join(HERE, "app/static/style.css"), encoding="utf-8").read()
logo_uri = "data:image/png;base64," + base64.b64encode(
    open(os.path.join(HERE, "app/static/darb_logo.png"), "rb").read()).decode()

def inline(html):
    html = html.replace('<link rel="stylesheet" href="/static/style.css">', "<style>\n" + css + "\n</style>")
    return html.replace("/static/darb_logo.png", logo_uri)

out = os.path.join(HERE, "..", "preview")
os.makedirs(out, exist_ok=True)
pages = {"dashboard": "/dashboard", "department": "/department/2", "evaluation": "/evaluation/1",
         "strategy": "/strategy", "report": "/report/print"}
for name, url in pages.items():
    r = c.get(url)
    open(os.path.join(out, name + ".html"), "w", encoding="utf-8").write(inline(r.text))
    print(name, r.status_code)
print("✅ previews in", os.path.abspath(out))
