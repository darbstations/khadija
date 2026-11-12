# -*- coding: utf-8 -*-
"""نسخةٌ للطباعة من الأوصاف الوظيفية — A4، صفحةٌ لكل دور

   المحتوى نفسُه الذي يبني به `build_jobs.py` ملفَّ إكسل — يُقرأ منه مباشرةً
   فلا تتفرّع نسختان.

   يُشغَّل من جذر المستودع:  PYTHONPATH=data python3 data/build_jobs_pdf.py
"""
import base64
import os
import subprocess

import build_jobs as J

HTML = "docs/الأوصاف-الوظيفية.html"
PDF = "docs/الأوصاف-الوظيفية.pdf"
JS = "data/jobs_pdf.js"
FONT = "data/fonts/noto-sans-arabic.woff2"

CSS = """
@font-face { font-family:"NSA"; src:url(data:font/woff2;base64,__FONT__) format("woff2");
             font-weight:100 900; font-style:normal; }
@page { size: A4 portrait; margin: 9mm 8mm; }
* { box-sizing: border-box; }
body { direction: rtl; margin: 0; color: #3D3D3D; font-family: "NSA", sans-serif; }
h1 { font-size: 14pt; margin: 0 0 0.8mm; color: #55565A; }
h1 small { font-size: 8pt; color: #9B968E; font-weight: 400; }
.sub { font-size: 7pt; color: #9B968E; margin: 0 0 2mm; }
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
.brk { page-break-before: always; }
.info td { font-size: 7.3pt; }
"""


def esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def role_html(name):
    j = J.JOBS[name]
    h = [f'<h1 class="brk">{esc(name)}  <small>— الإدارة التجارية</small></h1>',
         f'<div class="sub">{esc(j["scope"])} · يرفع إلى '
         f'{esc(j["reports"])} · يشرف على {esc(j["supervises"])}</div>']
    h.append('<h2>① الغرض من الوظيفة</h2><div class="purpose">'
             + "<br>".join(esc(x) for x in j["purpose"]) + "</div>")
    h.append('<h2>② الرقم الذي يُحاسَب عليه وحده</h2>'
             f'<div class="num">{esc(j["number"])}</div>')
    h.append("<h2>③ المسؤوليات</h2><table><tr><th>المسؤولية</th>"
             "<th>المخرج المطلوب</th><th>الدورية</th></tr>")
    for t, o, f in j["duties"]:
        h.append(f'<tr><td class="k">{esc(t)}</td><td>{esc(o)}</td>'
                 f'<td class="c">{esc(f)}</td></tr>')
    h.append("</table>")
    h.append("<h2>④ ما ليس من مهامه — ومالكُه</h2><table>"
             "<tr><th>البند</th><th>المالك</th></tr>")
    for t, o in j["nots"]:
        h.append(f'<tr><td class="no">{esc(t)}</td>'
                 f'<td class="own">{esc(o)}</td></tr>')
    h.append("</table>")
    h.append("<h2>⑤ حدود الصلاحية</h2><table><tr><th>البند</th><th>الحد</th>"
             "<th>من يعتمد ما يتجاوزه</th></tr>")
    for k, v, a in j["limits"]:
        cls = ' class="no"' if v.startswith(("ممنوع", "لا صلاحية", "غير مصرّح")) \
            else ' class="c"'
        h.append(f'<tr><td class="k">{esc(k)}</td><td{cls}>{esc(v)}</td>'
                 f'<td class="c">{esc(a)}</td></tr>')
    h.append("</table>")
    mine = [k for k in J.KPI if k[0] == name]
    h.append("<h2>⑥ مؤشراته</h2><table><tr><th>المؤشر</th><th>طريقة الحساب</th>"
             "<th>المصدر</th><th>الاتجاه</th><th>الوزن</th></tr>")
    for own, nm, calc, src, direc, wt in mine:
        h.append(f'<tr><td class="k">{esc(nm)}</td><td>{esc(calc)}</td>'
                 f'<td class="c">{esc(src)}</td><td class="c">{esc(direc)}</td>'
                 f'<td class="c">{wt*100:.0f}٪</td></tr>')
    h.append(f'<tr><td class="k">مجموع الأوزان</td><td colspan="3"></td>'
             f'<td class="c"><b>{sum(k[5] for k in mine)*100:.0f}٪</b></td></tr>')
    h.append("</table>")
    for title, pts, _ in j["rules"]:
        h.append(f'<div class="rule"><b>{esc(title)}</b>'
                 + "".join(f"<div>• {esc(p)}</div>" for p in pts) + "</div>")
    return "".join(h)


def build_html():
    css = CSS.replace("__FONT__",
                      base64.b64encode(open(FONT, "rb").read()).decode())
    H = ['<meta charset="utf-8"><style>' + css + "</style>"]
    H.append("<h1>الأوصاف الوظيفية — الإدارة التجارية</h1>")
    H.append('<div class="sub">ستة أدوار · لكل قرارٍ مالكٌ واحد · ولكل دورٍ '
             "رقمٌ لا يُقاس به غيره · و«ما ليس من مهامه» مكتوبٌ بمالكه</div>")
    H.append("<h2>خريطة المسؤولية</h2><table><tr><th>القرار أو الرقم</th>"
             "<th>المالك الوحيد</th><th>يُستشار</th><th>يُبلَّغ</th>"
             "<th>ليس له — والقاعدة</th></tr>")
    for q, own, cons, inf, no in J.MAP:
        H.append(f'<tr><td class="k">{esc(q)}</td>'
                 f'<td class="own">{esc(own)}</td>'
                 f'<td class="c">{esc(cons)}</td>'
                 f'<td class="c">{esc(inf)}</td>'
                 f'<td class="no">{esc(no)}</td></tr>')
    H.append("</table>")
    H.append('<div class="rule"><b>كيف تُقرأ</b>'
             "<div>• المالك الوحيد: من يُحاسَب على الرقم — لا من ينفّذه فقط</div>"
             "<div>• يُستشار: رأيه مطلوبٌ قبل القرار ولا يملك نقضه</div>"
             "<div>• ليس له: ما يُمنع صراحةً — وهو عمود منع الازدواجية</div>"
             "</div>")
    for name in J.ROLES:
        H.append(role_html(name))
    # صفحة المؤشرات
    H.append('<h1 class="brk">المؤشرات — لكل مؤشرٍ مالكٌ واحد</h1>')
    H.append('<div class="sub">٢٦ مؤشراً · لا مؤشر لمالكَين · ومجموع وزن كل '
             "مالكٍ ١٠٠٪</div><table><tr><th>المالك</th><th>المؤشر</th>"
             "<th>المصدر</th><th>الاتجاه</th><th>الوزن</th></tr>")
    for own, nm, calc, src, direc, wt in J.KPI:
        H.append(f'<tr><td class="own">{esc(own)}</td><td class="k">{esc(nm)}</td>'
                 f'<td class="c">{esc(src)}</td><td class="c">{esc(direc)}</td>'
                 f'<td class="c">{wt*100:.0f}٪</td></tr>')
    H.append("</table>")
    H.append('<div class="rule"><b>لماذا المدير يملك الريال والمشرف يملك العدد</b>'
             "<div>• تحليل موازنة ٢٠٢٧: الإيجار ٤٫٦٥× وحدةً × ٤٤٪ إيجاراً "
             "للوحدة — فالفجوة سعرُ الوحدة لا عددُها</div>"
             "<div>• ولو قيس الاثنان بالعدد لتنافسا على الرقم نفسه وتُرك "
             "السعر بلا مالك</div></div>")
    with open(HTML, "w", encoding="utf-8") as f:
        f.write("".join(H))
    return HTML


def build():
    build_html()
    with open(JS, "w", encoding="utf-8") as f:
        f.write("""const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch();
  const p = await b.newPage();
  await p.goto('file://' + process.argv[2], { waitUntil: 'networkidle' });
  await p.pdf({ path: process.argv[3], format: 'A4', printBackground: true });
  await b.close();
})();""")
    subprocess.run(["node", JS, os.path.abspath(HTML), os.path.abspath(PDF)],
                   check=True)
    os.remove(HTML)
    os.remove(JS)
    print(f"✓ {PDF}")


if __name__ == "__main__":
    build()
