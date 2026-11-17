# -*- coding: utf-8 -*-
"""مذكّرة داخلية بهوية درب — تُغلّف أي وصفٍ وظيفيٍّ بصيغةِ مراسلةٍ رسمية

   الترويسة أصلُها الرسميّ: استُخرجت طبقتاها من ملف الشركة
   (`التواصل الداخلي.docx`) وحُفظتا في `data/brand/`:
     · `darb-letterhead-logo.png`      — صفحةٌ بيضاء عليها الشعار أعلى اليمين
     · `darb-letterhead-watermark.jpg` — العلامة المائية «د» أسفل الصفحة
   وتُركَّبان كما يركّبهما وورد: المائيّة خلفاً، ثم الشعار، ثم النص فوقهما.

   ولونا الشعار الرسميان #F37521 و#59595C — وهما ما يُستعمل هنا.

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
LOGO = "data/brand/darb-letterhead-logo.png"
LOGO_ONLY = "data/brand/darb-logo.png"   # الشعار مقصوصاً — لرأس الطباعة
MARK = "data/brand/darb-letterhead-watermark.jpg"
ORANGE = "#F37521"       # من ملف الشعار الرسميّ
GREY = "#59595C"

HIJRI = {1: "محرّم", 2: "صفر", 3: "ربيع الأول", 4: "ربيع الآخر",
         5: "جمادى الأولى", 6: "جمادى الآخرة", 7: "رجب", 8: "شعبان",
         9: "رمضان", 10: "شوّال", 11: "ذو القعدة", 12: "ذو الحجّة"}

CSS = """
@font-face { font-family:"NSA"; src:url(data:font/woff2;base64,__FONT__) format("woff2");
             font-weight:100 900; font-style:normal; }
@page { size: A4 portrait; }
* { box-sizing: border-box; }
html { background-image: url(__MARK__); background-repeat: no-repeat;
       background-position: left bottom; background-size: 100% 100%;
       background-attachment: fixed; }
body { direction: rtl; margin: 0; color: #3D3D3D; font-family: "NSA", sans-serif; }

/* ── طبقتا الترويسة الرسمية — تتكرّران على كل صفحة */
/* الشعار وسطرُ الذيل يتكرّران عبر رأس/ذيل الطباعة في كروميوم.
   والعلامة المائية خلفيةٌ على الصفحة نفسها. */
.page { position: relative; z-index: 1; }

/* ── شريط العنوان تحت الشعار */
.head { border-bottom: 1.2pt solid __ORANGE__; padding-bottom: 2.5mm;
        margin: 0 0 3mm; display: flex; align-items: flex-end; }
.org { flex: 1; }
.org .n { font-size: 12pt; font-weight: 700; color: __GREY__; }
.org .d { font-size: 7.5pt; color: #9B968E; margin-top: 0.6mm; }
.kind { font-size: 8.5pt; color: __ORANGE__; font-weight: 700; white-space: nowrap; }

/* ── حقول المراسلة */
table.meta { width: 100%; border-collapse: collapse; font-size: 7.4pt;
             margin-bottom: 3mm; }
table.meta td { border: 0.4pt solid #E3DCD1; padding: 1.1mm 2mm; }
table.meta td.k { background: #F5F2ED; font-weight: 700; color: __GREY__;
                  white-space: nowrap; width: 20mm; }

/* ── الجسم */
h1 { font-size: 12pt; margin: 0 0 1mm; color: __GREY__; }
.sub { font-size: 7pt; color: #9B968E; margin: 0 0 2.5mm; }
h2 { font-size: 8.5pt; color: __ORANGE__; background: #FBEEE0; margin: 2.2mm 0 0.9mm;
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
.sig th { background: #F5F2ED; color: __GREY__; border: 0.4pt solid #E3DCD1;
          padding: 1.2mm; font-size: 7pt; }
.sig td { height: 15mm; border: 0.4pt solid #E3DCD1; vertical-align: bottom;
          padding: 1.2mm; color: #B5AEA2; font-size: 6pt; }
.ack { margin-top: 2.5mm; border: 0.5pt solid __ORANGE__; border-radius: 1mm;
       padding: 2mm; font-size: 7pt; color: #55565A; }
.ack b { color: __ORANGE__; }
.ack { color: __GREY__; }
.ack .line { display: inline-block; border-bottom: 0.5pt dotted #9B968E;
             min-width: 38mm; margin: 0 1mm; }
.foot { margin-top: 3mm; padding-top: 1.5mm; text-align: center;
        font-size: 5.8pt; color: #B5AEA2; border-top: 0.4pt solid #ECE6DD; }
"""


def memo_html(role, ref, to, frm, note=""):
    b64 = lambda p: base64.b64encode(open(p, "rb").read()).decode()
    css = (CSS.replace("__FONT__", b64(FONT))
              .replace("__ORANGE__", ORANGE).replace("__GREY__", GREY)
              .replace("__MARK__", "data:image/jpeg;base64," + b64(MARK)))
    today = dt.date.today()
    j = J.JOBS[role]
    H = ['<meta charset="utf-8"><style>' + css + "</style>"]
    # طبقتا الترويسة الرسمية — كما يركّبهما وورد: المائيّة ثم الشعار
    H.append('<div class="page">')
    H.append(
        '<div class="head">'
        '<div class="org"><div class="n">شركة درب الوقود المحدودة</div>'
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
    H.append('</div>')
    path = f"docs/مذكّرة-وصف-وظيفي-{role.replace(' ', '-')}.html"
    with open(path, "w", encoding="utf-8") as f:
        f.write("".join(H))
    return path, path.replace(".html", ".pdf")


def build(role, ref, to, frm, note=""):
    html, pdf = memo_html(role, ref, to, frm, note)
    logo64 = base64.b64encode(open(LOGO_ONLY, "rb").read()).decode()
    with open(JS, "w", encoding="utf-8") as f:
        f.write("""const { chromium } = require('playwright');
const fs = require('fs');
(async () => {
  const b = await chromium.launch();
  const p = await b.newPage();
  await p.goto('file://' + process.argv[2], { waitUntil: 'networkidle' });
  await p.pdf({
    path: process.argv[3], format: 'A4', printBackground: true,
    displayHeaderFooter: true,
    margin: { top: '27mm', right: '13mm', bottom: '16mm', left: '13mm' },
    headerTemplate: '<div style="width:100%;padding:9mm 13mm 0;">'
      + '<img src="' + process.argv[4] + '" style="height:12mm">' + '</div>',
    footerTemplate: '<div style="width:100%;text-align:center;font-size:6pt;'
      + 'color:#B5AEA2;font-family:sans-serif;padding-bottom:6mm;">'
      + 'وثيقة داخلية — تُحفظ في ملف الموظف ولا تُتداول خارج الشركة'
      + ' &nbsp;·&nbsp; <span class="pageNumber"></span>/'
      + '<span class="totalPages"></span></div>',
  });
  await b.close();
})();""")
    subprocess.run(["node", JS, os.path.abspath(html), os.path.abspath(pdf),
                    "data:image/png;base64," + logo64], check=True)
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
