# -*- coding: utf-8 -*-
"""مذكّرة داخلية بهوية درب — تُغلّف أي وصفٍ وظيفيٍّ بصيغةِ مراسلةٍ رسمية

   لا يوجد في المشروع ملفُّ ترويسةٍ رسميٍّ ولا شعار. والموجود هوية درب
   البصرية وحدها (`data/brand.py` — مستخرجةٌ من تقارير الشركة). فهذا
   الملف يبني المذكّرة بتلك الهوية ويترك موضع الشعار فارغاً معلَّماً،
   يُستبدل بالشعار حين يصل.

   وحقول المراسلة الرسمية كاملة: الرقم المرجعي · التاريخان · المرسل
   والمرسَل إليه · الموضوع · المرفقات · وبيوت التوقيع والاستلام.

   يُشغَّل:  PYTHONPATH=data python3 data/build_memo.py "مشرف العقار"
"""
import base64
import os
import subprocess
import sys
import datetime as dt

import build_jobs as J
import build_jobs_pdf as P

JS = "data/memo.js"
FONT = "data/fonts/noto-sans-arabic.woff2"

HIJRI = {1: "محرّم", 2: "صفر", 3: "ربيع الأول", 4: "ربيع الآخر",
         5: "جمادى الأولى", 6: "جمادى الآخرة", 7: "رجب", 8: "شعبان",
         9: "رمضان", 10: "شوّال", 11: "ذو القعدة", 12: "ذو الحجّة"}

CSS = """
@font-face { font-family:"NSA"; src:url(data:font/woff2;base64,__FONT__) format("woff2");
             font-weight:100 900; font-style:normal; }
@page { size: A4 portrait; margin: 10mm 9mm 11mm; }
* { box-sizing: border-box; }
body { direction: rtl; margin: 0; color: #3D3D3D; font-family: "NSA", sans-serif; }

/* ── الترويسة */
.head { display: flex; align-items: center; gap: 4mm; border-bottom: 1.2pt solid #F5831F;
        padding-bottom: 2.5mm; margin-bottom: 3mm; }
.logo { width: 26mm; height: 14mm; border: 0.6pt dashed #C9C2B6; border-radius: 1mm;
        display: flex; align-items: center; justify-content: center;
        font-size: 5.6pt; color: #B5AEA2; text-align: center; line-height: 1.3; }
.org { flex: 1; }
.org .n { font-size: 13pt; font-weight: 700; color: #55565A; }
.org .d { font-size: 7.5pt; color: #9B968E; margin-top: 0.6mm; }
.kind { font-size: 8pt; color: #F5831F; font-weight: 700; text-align: left;
        white-space: nowrap; }

/* ── حقول المراسلة */
table.meta { width: 100%; border-collapse: collapse; font-size: 7.4pt;
             margin-bottom: 3mm; }
table.meta td { border: 0.4pt solid #E3DCD1; padding: 1.1mm 2mm; }
table.meta td.k { background: #F5F2ED; font-weight: 700; color: #55565A;
                  white-space: nowrap; width: 20mm; }

/* ── الجسم */
h1 { font-size: 12pt; margin: 0 0 1mm; color: #55565A; }
.sub { font-size: 7pt; color: #9B968E; margin: 0 0 2.5mm; }
h2 { font-size: 8.5pt; color: #F5831F; background: #FBEEE0; margin: 2.2mm 0 0.9mm;
     padding: 0.8mm 2mm; border-radius: 1mm; page-break-after: avoid; }
table { border-collapse: collapse; width: 100%; font-size: 6.6pt;
        page-break-inside: avoid; }
th { background: #55565A; color: #fff; font-weight: 600; padding: 0.8mm 1mm;
     font-size: 6.3pt; text-align: center; }
td { border: 0.35pt solid #E3DCD1; padding: 0.6mm 1mm; vertical-align: top;
     text-align: right; }
td.c { text-align: center; white-space: nowrap; }
td.k { font-weight: 700; background: #FAF8F4; white-space: nowrap; }
.num { background: #D6E9E1; color: #1F5E4A; font-weight: 700; font-size: 9pt;
       text-align: center; padding: 1.4mm; border-radius: 1mm; }
.purpose { font-size: 7.8pt; line-height: 1.6; }
td.no { color: #9A3E2C; background: #F2DAD4; font-weight: 600; }
td.own { color: #2E8B6F; font-weight: 700; white-space: nowrap; }
.rule { background: #F2DAD4; color: #9A3E2C; font-size: 6.4pt; padding: 1mm 2mm;
        margin-top: 1mm; border-radius: 1mm; }
.rule b { display: block; margin-bottom: 0.6mm; }
.rule div { margin: 0.4mm 0; }

/* ── التواقيع */
.sig { width: 100%; border-collapse: collapse; font-size: 7pt; margin-top: 4mm;
       page-break-inside: avoid; }
.sig th { background: #F5F2ED; color: #55565A; border: 0.4pt solid #E3DCD1;
          padding: 1.2mm; font-size: 7pt; }
.sig td { height: 15mm; border: 0.4pt solid #E3DCD1; vertical-align: bottom;
          padding: 1.2mm; color: #B5AEA2; font-size: 6pt; }
.ack { margin-top: 2.5mm; border: 0.5pt solid #F5831F; border-radius: 1mm;
       padding: 2mm; font-size: 7pt; color: #55565A; }
.ack b { color: #F5831F; }
.ack .line { display: inline-block; border-bottom: 0.5pt dotted #9B968E;
             min-width: 38mm; margin: 0 1mm; }
.foot { margin-top: 3mm; padding-top: 1.5mm; text-align: center;
        font-size: 5.8pt; color: #B5AEA2; border-top: 0.4pt solid #ECE6DD; }
"""


def memo_html(role, ref, to, frm, note=""):
    css = CSS.replace("__FONT__",
                      base64.b64encode(open(FONT, "rb").read()).decode())
    today = dt.date.today()
    j = J.JOBS[role]
    H = ['<meta charset="utf-8"><style>' + css + "</style>"]
    # الترويسة
    H.append(
        '<div class="head">'
        '<div class="logo">موضع الشعار<br>الرسمي</div>'
        '<div class="org"><div class="n">شركة درب لمحطات الوقود</div>'
        '<div class="d">القطاع التجاري · إدارة العقار</div></div>'
        '<div class="kind">مذكّرة داخلية</div></div>')
    # الحقول
    H.append(
        '<table class="meta">'
        f'<tr><td class="k">الرقم المرجعي</td><td>{P.esc(ref)}</td>'
        f'<td class="k">التاريخ</td><td>{today.strftime("%Y/%m/%d")}م</td></tr>'
        f'<tr><td class="k">إلى</td><td>{P.esc(to)}</td>'
        f'<td class="k">من</td><td>{P.esc(frm)}</td></tr>'
        f'<tr><td class="k">الموضوع</td><td colspan="3">'
        f'اعتماد الوصف الوظيفي — {P.esc(role)}</td></tr>'
        '<tr><td class="k">المرفقات</td><td colspan="3">'
        'خريطة المسؤولية · جدول المؤشرات</td></tr></table>')
    if note:
        H.append(f'<div class="purpose" style="margin-bottom:2.5mm">{P.esc(note)}</div>')
    # الجسم — من المصدر نفسه
    body = P.role_html(role).replace('<h1 class="brk">', "<h1>", 1)
    H.append(body)
    # التواقيع
    H.append(
        '<table class="sig"><tr><th>أعدّها</th><th>راجعها</th>'
        '<th>اعتمدها</th></tr>'
        f'<tr><td>مدير العقار<br>الاسم والتوقيع</td>'
        f'<td>{P.esc(J.HEAD)}<br>الاسم والتوقيع</td>'
        '<td>الموارد البشرية<br>الاسم والتوقيع</td></tr></table>')
    H.append(
        '<div class="ack"><b>إقرار استلام</b> — اطّلعتُ على هذا الوصف وما فيه '
        'من مسؤولياتٍ وحدودِ صلاحيةٍ ومؤشرات، وتسلّمتُ نسخةً منه. '
        'الاسم <span class="line"></span> التوقيع <span class="line"></span> '
        'التاريخ <span class="line"></span></div>')
    H.append('<div class="foot">وثيقة داخلية — تُحفظ في ملف الموظف ولا تُتداول '
             'خارج الشركة</div>')
    path = f"docs/مذكّرة-وصف-وظيفي-{role.replace(' ', '-')}.html"
    with open(path, "w", encoding="utf-8") as f:
        f.write("".join(H))
    return path, path.replace(".html", ".pdf")


def build(role, ref, to, frm, note=""):
    html, pdf = memo_html(role, ref, to, frm, note)
    with open(JS, "w", encoding="utf-8") as f:
        f.write("""const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch();
  const p = await b.newPage();
  await p.goto('file://' + process.argv[2], { waitUntil: 'networkidle' });
  await p.pdf({ path: process.argv[3], format: 'A4', printBackground: true });
  await b.close();
})();""")
    subprocess.run(["node", JS, os.path.abspath(html), os.path.abspath(pdf)],
                   check=True)
    os.remove(html)
    os.remove(JS)
    print(f"✓ {pdf}")


if __name__ == "__main__":
    role = sys.argv[1] if len(sys.argv) > 1 else "مشرف العقار"
    if role not in J.JOBS:
        sys.exit("الأدوار: " + " · ".join(J.ROLES))
    build(role,
          ref=sys.argv[2] if len(sys.argv) > 2 else "ق‌ت/ع/٢٠٢٦/٠٠١",
          to="مشرف العقار",
          frm="مدير العقار — القطاع التجاري",
          note="إلحاقاً بتنظيم مسؤوليات إدارة العقار، نرفق لكم الوصف "
               "الوظيفي المعتمد للدور، موضّحاً الرقم الذي يُقاس به الأداء، "
               "والمسؤوليات ومخرجاتها، وما هو خارج نطاق الدور ومالكه، "
               "وحدود الصلاحية، والمؤشرات وأوزانها. "
               "ويُعمل به من تاريخ اعتماده.")
