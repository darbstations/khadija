# -*- coding: utf-8 -*-
"""خطابٌ رسميٌّ على ترويسة درب — يُغلّف أي وصفٍ وظيفيٍّ بصيغةِ مراسلة

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
FONT = "data/fonts/din-next-lt-arabic-medium.ttf"   # خط درب الرسميّ من الترويسة
LOGO = "data/brand/darb-letterhead-logo.png"
LOGO_ONLY = "data/brand/darb-logo.png"   # الشعار مقصوصاً — لرأس الطباعة
MARK = "data/brand/darb-letterhead-watermark.jpg"
ORANGE = "#F37521"       # من ملف الشعار الرسميّ
GREY = "#59595C"

HIJRI = {1: "محرّم", 2: "صفر", 3: "ربيع الأول", 4: "ربيع الآخر",
         5: "جمادى الأولى", 6: "جمادى الآخرة", 7: "رجب", 8: "شعبان",
         9: "رمضان", 10: "شوّال", 11: "ذو القعدة", 12: "ذو الحجّة"}

CSS = """
@font-face { font-family:"DIN"; src:url(data:font/ttf;base64,__FONT__) format("truetype");
             font-weight:100 900; font-style:normal; }
@page { size: A4 portrait; }
* { box-sizing: border-box; }
html { background-image: url(__MARK__); background-repeat: no-repeat;
       background-position: left bottom; background-size: 100% 100%;
       background-attachment: fixed; }
body { direction: rtl; margin: 0; color: #3D3D3D; font-family: "DIN", sans-serif; }
.page { position: relative; z-index: 1; }

/* ── شريط العنوان */
.head { border-bottom: 1pt solid __ORANGE__; padding-bottom: 2.5mm;
        margin: 0 0 3mm; }
.head .n { font-size: 12pt; color: __GREY__; }
.head .d { font-size: 8pt; color: #8E8A84; margin-top: 0.8mm; }

/* ── حقول المراسلة */
table.meta { width: 100%; border-collapse: collapse; font-size: 8pt;
             margin-bottom: 3mm; }
table.meta td { border: 0.4pt solid #E6E2DB; padding: 1.3mm 2mm; }
table.meta td.k { background: #F5F2ED; color: __GREY__; white-space: nowrap;
                  width: 20mm; }

/* ── الجسم */
h1 { font-size: 12.5pt; margin: 0 0 1mm; color: __GREY__; }
.sub { font-size: 7.5pt; color: #8E8A84; margin: 0 0 3mm; }
h2 { font-size: 9pt; color: __GREY__; background: #F5F2ED; margin: 3mm 0 1mm;
     padding: 1mm 2mm; page-break-after: avoid;
     border-right: 1.6pt solid __ORANGE__; }
table { border-collapse: collapse; width: 100%; font-size: 7.2pt;
        page-break-inside: avoid; }
th { background: __GREY__; color: #fff; padding: 1mm 1mm; font-size: 6.8pt;
     text-align: center; }
td { border: 0.35pt solid #E6E2DB; padding: 0.9mm 1.2mm; vertical-align: top;
     text-align: right; }
td.c { text-align: center; white-space: nowrap; }
td.k { background: #FAF9F7; color: __GREY__; white-space: nowrap; }
.num { background: #F5F2ED; color: __GREY__; font-size: 10pt;
       text-align: center; padding: 2mm; border-right: 1.6pt solid __ORANGE__; }
.purpose { font-size: 8.2pt; line-height: 1.75; }
td.no { background: #FAF9F7; color: __GREY__; }
td.own { color: __ORANGE__; white-space: nowrap; }
.rule { background: #F5F2ED; color: #6E6A64; font-size: 7pt; padding: 1.4mm 2mm;
        margin-top: 1.2mm; border-right: 1.6pt solid __ORANGE__; }
.rule b { display: block; margin-bottom: 0.8mm; color: __GREY__; }
.rule div { margin: 0.5mm 0; }

/* ── التواقيع */
.sig { width: 100%; border-collapse: collapse; font-size: 7.4pt; margin-top: 5mm;
       page-break-inside: avoid; }
.sig th { background: #F5F2ED; color: __GREY__; border: 0.4pt solid #E6E2DB;
          padding: 1.4mm; font-size: 7.4pt; }
.sig td { height: 16mm; border: 0.4pt solid #E6E2DB; vertical-align: bottom;
          padding: 1.4mm; color: #A8A49E; font-size: 6.4pt; }
.ack { margin-top: 3mm; border: 0.5pt solid __ORANGE__; padding: 2.2mm;
       font-size: 7.4pt; color: __GREY__; }
.ack b { color: __ORANGE__; }
.ack .line { display: inline-block; border-bottom: 0.5pt dotted #A8A49E;
             min-width: 38mm; margin: 0 1mm; }
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
        '<div class="n">شركة درب الوقود المحدودة</div>'
        '<div class="d">القطاع التجاري · إدارة العقار</div></div>')
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
      + 'تُحفظ في ملف الموظف'
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
