# -*- coding: utf-8 -*-
"""نسخةٌ للطباعة من جدول المقارنة — صفحةُ A4 واحدةٌ أو اثنتان، أرقامٌ بلا شرح

   المصدر والأرقام هي نفسُها التي يبني بها `build_budget_compare.py` ملفَّ
   إكسل — فلا يتخلّف رقمٌ عن نظيره حين يتغيّر المصدر.

   يُشغَّل من جذر المستودع:
       python3 data/build_budget_compare_pdf.py        # يُخرج HTML
       node  data/pdf.js  <html>  <pdf>                # ثم يطبعه
   أو يُشغَّل `make_pdf()` فيفعل الاثنين.
"""
import json
import os
import subprocess
import sys

SRC = "outlets/data/budget-2027.json"
HTML = "docs/مقارنة-المخطط-والفعلي-والموازنة.html"
PDF = "docs/مقارنة-المخطط-والفعلي-والموازنة.pdf"
JS = "data/pdf.js"
# خطٌّ عربيٌّ مُضمَّن — فلا خطَّ عربيّ في النظام، وبدونه تنفصل الحروف
FONT = "data/fonts/noto-sans-arabic.woff2"

M = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
     "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"]
D26 = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

CSS = """
@font-face { font-family: "NSA"; src: url(data:font/woff2;base64,__FONT__) format("woff2");
             font-weight: 100 900; font-style: normal; }
@page { size: A4 portrait; margin: 8mm 7mm; }
* { box-sizing: border-box; }
body { direction: rtl; margin: 0; color: #3D3D3D;
       font-family: "NSA", "FreeSans", sans-serif; }
h1 { font-size: 15pt; margin: 0 0 0.8mm; color: #55565A; }
.sub { font-size: 7pt; color: #9B968E; margin: 0 0 2.5mm; }
h2 { font-size: 9pt; color: #F5831F; background: #FBEEE0; margin: 2.4mm 0 0.8mm;
     padding: 1mm 2mm; border-radius: 1mm; }
table, .n, h2 { page-break-inside: avoid; }
h2 { page-break-after: avoid; }
table { border-collapse: collapse; width: 100%; font-size: 6.2pt; }
th { background: #55565A; color: #fff; font-weight: 600; padding: 0.9mm 0.7mm;
     text-align: center; font-size: 6.5pt; }
td { border: 0.35pt solid #E3DCD1; padding: 0.5mm 0.7mm; text-align: center;
     white-space: nowrap; }
td.k { text-align: right; font-weight: 700; background: #FAF8F4; }
tr.t td { background: #EDE7DE; font-weight: 700; }
td.fc { background: #FCE7C8; }
td.c { background: #F5F2ED; color: #6E6A64; }
.n { font-size: 6.6pt; padding: 1mm 2mm; margin: 1mm 0 0; border-radius: 1mm; }
.gold { background: #FCE7C8; color: #A8681A; }
.bad { background: #F2DAD4; color: #9A3E2C; }
.ask { font-size: 6.8pt; line-height: 1.7; margin: 1.2mm 0 0; }
.ask b { color: #55565A; }
"""


def ar(x, dec=0):
    if x is None or x == "":
        return "—"
    if isinstance(x, str):
        return x
    return f"{x:,.{dec}f}"


def pc(x):
    return "—" if x is None else f"{x*100:.0f}٪"


def build_html():
    d = json.load(open(SRC, encoding="utf-8"))
    ra, rr = d["fuel_actual_months"], d["rent_actual_months"]

    def lpd(lit, cnt, upto=12):
        sd = sum(cnt[k] * D26[k] for k in range(upto))
        return sum(lit[:upto]) / sd if sd else 0

    def per(b, k):
        return d[b][k] / d[b + "_n"][k] if d[b + "_n"][k] else 0

    import base64
    css = CSS.replace("__FONT__",
                      base64.b64encode(open(FONT, "rb").read()).decode())
    H = ['<meta charset="utf-8"><style>' + css + "</style>",
         "<h1>المخطط · الفعلي · الموازنة</h1>",
         f'<div class="sub">المصدر: {d["source"]} — الفعلي: الوقود تسعةُ '
         f"أشهر، والإيجار ثمانية. والمظلَّل ▸ منقولٌ لا مقيس.</div>"]

    # ── ⓪ الملخّص
    H.append("<h2>⓪ الملخّص · المحطات والوقود والوحدات</h2><table>")
    H.append("<tr><th>البند</th><th>المخطط</th><th>الفعلي</th>"
             "<th>الموازنة ٢٧</th><th>الفعلي ÷ المخطط</th><th>٢٧ ÷ الفعلي</th></tr>")
    rows = [
        ("محطةٌ قائمةٌ بنهاية السنة",
         d["register_existing"] + d["register_opens_26"],
         d["fuel_l26_n"][ra - 1], d["fuel_has_27"], 0),
        ("محطةٌ باعت في أي شهر", None, d["fuel_sold_actual"], d["fuel_has_27"], 0),
        ("الوقود · مليون لتر", None, sum(d["fuel_l26"][:ra]) / 1e6,
         sum(d["fuel_l27"]) / 1e6, 0),
        ("الوقود · مليون ريال", None, sum(d["fuel_v26"][:ra]) / 1e6,
         sum(d["fuel_v27"]) / 1e6, 0),
        ("لتر / محطة / يوم", None, lpd(d["fuel_l26"], d["fuel_l26_n"], ra),
         lpd(d["fuel_l27"], d["fuel_l27_n"]), 0),
        ("وحدةٌ تفوتر · مخطط «أ»", d["rent_p26a_n"][11],
         d["rent_gross_n"][rr - 1], d["rent_b27_n"][11], 0),
        ("وحدةٌ تفوتر · مخطط «ب»", d["rent_p26b_n"][11],
         d["rent_gross_n"][rr - 1], d["rent_b27_n"][11], 0),
        ("الإيجار · مليون ريال «أ»", sum(d["rent_p26a"]) / 1e6,
         sum(d["rent_gross"]) / 1e6, sum(d["rent_b27"]) / 1e6, 1),
        ("الإيجار · مليون ريال «ب»", sum(d["rent_p26b"]) / 1e6,
         sum(d["rent_gross"]) / 1e6, sum(d["rent_b27"]) / 1e6, 1),
        ("ريال / وحدة / شهر", per("rent_p26b", 11), per("rent_gross", rr - 1),
         per("rent_b27", 11), 0),
    ]
    for lbl, p, a, b, dec in rows:
        rt = pc(a / p) if p else "—"
        gw = f"{b/a:,.1f}×" if a else "—"
        H.append(f'<tr><td class="k">{lbl}</td><td>{ar(p, dec)}</td>'
                 f"<td>{ar(a, dec)}</td><td>{ar(b, dec)}</td>"
                 f'<td class="c">{rt}</td><td class="c">{gw}</td></tr>')
    H.append("</table>")
    H.append(f'<div class="n gold">مصدرُ النمو — الوقود: '
             f'{d["fuel_has_27"]/d["fuel_l26_n"][ra-1]:.2f}× محطة × '
             f'{lpd(d["fuel_l27"],d["fuel_l27_n"])/lpd(d["fuel_l26"],d["fuel_l26_n"],ra)*100:.0f}٪ '
             f"إنتاجية. &nbsp; الإيجار: "
             f'{d["rent_b27_n"][11]/d["rent_gross_n"][rr-1]:.2f}× وحدة × '
             f'{per("rent_b27",11)/per("rent_gross",rr-1)*100:.0f}٪ إيجاراً.</div>')

    # ── ① الإيجار · الإيراد
    H.append("<h2>① الإيجار · الإيراد المفوتَر (ريال)</h2><table>")
    H.append("<tr><th>الشهر</th><th>مخطط ‹أ›</th><th>مخطط ‹ب›</th><th>الفعلي</th>"
             "<th>÷أ</th><th>÷ب</th><th>موازنة ٢٧</th><th>٢٧÷الفعلي</th></tr>")
    for k in range(12):
        fc = ' class="fc"' if k >= rr else ""
        a, b2, ac = d["rent_p26a"][k], d["rent_p26b"][k], d["rent_gross"][k]
        H.append(f'<tr><td class="k">{M[k]}{" ▸" if k>=rr else ""}</td>'
                 f"<td>{ar(a)}</td><td>{ar(b2)}</td><td{fc}>{ar(ac)}</td>"
                 f'<td class="c">{pc(ac/a) if a else "—"}</td>'
                 f'<td class="c">{pc(ac/b2) if b2 else "—"}</td>'
                 f'<td>{ar(d["rent_b27"][k])}</td>'
                 f'<td class="c">{d["rent_b27"][k]/ac:,.1f}×</td></tr>')
    sa, sb, sc = sum(d["rent_p26a"]), sum(d["rent_p26b"]), sum(d["rent_gross"])
    H.append(f'<tr class="t"><td class="k">السنة</td><td>{ar(sa)}</td>'
             f"<td>{ar(sb)}</td><td>{ar(sc)}</td><td>{pc(sc/sa)}</td>"
             f'<td>{pc(sc/sb)}</td><td>{ar(sum(d["rent_b27"]))}</td>'
             f'<td>{sum(d["rent_b27"])/sc:,.1f}×</td></tr>')
    pa = sum(d["rent_p26a"][:rr]); pb = sum(d["rent_p26b"][:rr])
    pc_ = sum(d["rent_gross"][:rr])
    H.append(f'<tr class="t"><td class="k">الفعلي المعتمد ‹ينا–أغس›</td>'
             f"<td>{ar(pa)}</td><td>{ar(pb)}</td><td>{ar(pc_)}</td>"
             f"<td>{pc(pc_/pa)}</td><td>{pc(pc_/pb)}</td><td>—</td>"
             f"<td>—</td></tr></table>")
    H.append('<div class="n bad">«أ»: وحدات ٢٨٪ وإيراد ٩٠٪. &nbsp; '
             "«ب»: وحدات ١٤٦٪ وإيراد ٩٩٪ — الفجوة سعرُ الوحدة لا عددُها.</div>")

    # ── ② الإيجار · الوحدات
    H.append('<h2>② الإيجار · عدد الوحدات المفوتَرة</h2><table>')
    H.append("<tr><th>الشهر</th><th>مخطط ‹أ›</th><th>مخطط ‹ب›</th><th>الفعلي</th>"
             "<th>÷أ</th><th>÷ب</th><th>موازنة ٢٧</th><th>٢٧÷الفعلي</th></tr>")
    for k in range(12):
        fc = ' class="fc"' if k >= rr else ""
        a, b2, ac = d["rent_p26a_n"][k], d["rent_p26b_n"][k], d["rent_gross_n"][k]
        H.append(f'<tr><td class="k">{M[k]}{" ▸" if k>=rr else ""}</td>'
                 f"<td>{ar(a)}</td><td>{ar(b2)}</td><td{fc}>{ar(ac)}</td>"
                 f'<td class="c">{pc(ac/a)}</td><td class="c">{pc(ac/b2)}</td>'
                 f'<td>{ar(d["rent_b27_n"][k])}</td>'
                 f'<td class="c">{d["rent_b27_n"][k]/ac:,.1f}×</td></tr>')
    pace = lambda s, last: (s[last] - s[0]) / last
    H.append(f'<tr class="t"><td class="k">صافي الزيادة شهرياً</td>'
             f'<td>{pace(d["rent_p26a_n"],11):,.1f}</td>'
             f'<td>{pace(d["rent_p26b_n"],11):,.1f}</td>'
             f'<td>{pace(d["rent_gross_n"],rr-1):,.1f}</td><td>—</td><td>—</td>'
             f'<td>{pace(d["rent_b27_n"],11):,.1f}</td><td>—</td></tr></table>')
    H.append(f'<div class="n bad">‎+{d["rent_b27_n"][0]-d["rent_gross_n"][rr-1]} '
             f"وحدة قبل يناير ٢٧، ثم "
             f'‎+{pace(d["rent_b27_n"],11):.0f} شهرياً — مقابل '
             f'‎+{pace(d["rent_gross_n"],rr-1):.1f} محقَّقة = '
             f'{pace(d["rent_b27_n"],11)/pace(d["rent_gross_n"],rr-1):.1f}×. &nbsp; '
             f'والشاغر في موازنة ٢٧ ({d["rent_b27_vacant_notactive"]} وحدة) '
             f"كلُّه في محطاتٍ لم تُفتح، والقابل للتأجير اليوم "
             f'({d["rent_pool_active_vacant"]} وحدة) موازنتُه صفر.</div>')

    # ── ③ الوقود
    H.append("<h2>③ الوقود · اللترات والمحطات</h2><table>")
    H.append("<tr><th>الشهر</th><th>فعلي · لتر</th><th>محطة</th><th>ريال/لتر</th>"
             "<th>موازنة ٢٧ · لتر</th><th>محطة</th><th>ريال/لتر</th>"
             "<th>٢٧÷٢٦</th></tr>")
    for k in range(12):
        fc = ' class="fc"' if k >= ra else ""
        l6, n6 = d["fuel_l26"][k], d["fuel_l26_n"][k]
        l7, n7 = d["fuel_l27"][k], d["fuel_l27_n"][k]
        H.append(f'<tr><td class="k">{M[k]}{" ▸" if k>=ra else ""}</td>'
                 f"<td{fc}>{ar(l6)}</td><td{fc}>{ar(n6)}</td>"
                 f'<td class="c">{d["fuel_v26"][k]/l6:,.3f}</td>'
                 f"<td>{ar(l7)}</td><td>{ar(n7)}</td>"
                 f'<td class="c">{d["fuel_v27"][k]/l7:,.3f}</td>'
                 f'<td class="c">{l7/l6:,.1f}×</td></tr>')
    t6, t7 = sum(d["fuel_l26"]), sum(d["fuel_l27"])
    H.append(f'<tr class="t"><td class="k">السنة</td><td>{ar(t6)}</td>'
             f'<td>{ar(d["fuel_any_26"])}</td>'
             f'<td>{sum(d["fuel_v26"])/t6:,.3f}</td><td>{ar(t7)}</td>'
             f'<td>{ar(d["fuel_has_27"])}</td>'
             f'<td>{sum(d["fuel_v27"])/t7:,.3f}</td>'
             f"<td>{t7/t6:,.1f}×</td></tr>")
    a6 = sum(d["fuel_l26"][:ra])
    H.append(f'<tr class="t"><td class="k">الفعلي ‹ينا–سبت›</td><td>{ar(a6)}</td>'
             f'<td>{ar(d["fuel_sold_actual"])}</td>'
             f'<td>{sum(d["fuel_v26"][:ra])/a6:,.3f}</td>'
             f"<td colspan=4>حصتُه من ٢٦: {pc(a6/t6)}</td></tr></table>")
    H.append(f'<div class="n bad">لا مخطَّط وقودٍ لـ٢٦ في الملف (عمودا P-2026 '
             f"صفر). &nbsp; ومحطات ٢٧ الجديدة "
             f'({d["fuel_rate_new_27"]["n"]}): وسيط '
             f'{d["fuel_rate_new_27"]["median"]:,.0f} لتر/يوم وأدناها '
             f'{d["fuel_rate_new_27"]["min"]:,.0f} — مقابل '
             f'{d["fuel_rate_old_27"]["median"]:,.0f} للقائمة.</div>')

    # ── ④ المحطات + ⑤ الناقص
    H.append("<h2>④ المحطات · السجل مقابل ما باع فعلاً</h2><table>")
    H.append("<tr><th>البند</th><th>عدد</th><th>البند</th><th>عدد</th></tr>")
    pairs = [("سجلُّ المحطات", d["register_total"],
              "تفتح نوفمبر–ديسمبر ٢٦", d["register_opens_26"]),
             ("قائمة ‹Existing›", d["register_existing"],
              "تفتح خلال ٢٠٢٧", d["register_opens_27"]),
             ("باعت فعلاً ينا–سبت", d["fuel_sold_actual"],
              "لها خطةُ لترٍ في ٢٧", d["fuel_has_27"]),
             ("أعلى شهرٍ فعليّ ‹سبتمبر›", d["fuel_l26_n"][ra - 1],
              "جديدةٌ كلياً في ٢٧", d["fuel_new_27"])]
    for a, b, c, e in pairs:
        H.append(f'<tr><td class="k">{a}</td><td>{ar(b)}</td>'
                 f'<td class="k">{c}</td><td>{ar(e)}</td></tr>')
    H.append("</table>")
    H.append(f'<div class="n gold">'
             f'{d["register_existing"]-d["fuel_sold_actual"]} محطةً «قائمة» '
             f"بلا لترٍ واحد في تسعة أشهر.</div>")

    H.append("<h2>⑤ ما يُطلَب قبل الاعتماد</h2>")
    H.append('<div class="ask">'
             "<b>المالية:</b> مخطَّط وقود ٢٠٢٦ (غيرُ موجود) · أيُّ نسختي مخطط "
             f"الإيجار معتمدة · سببُ إشعارات أغسطس "
             f'({d["rent_credit"][7]/1e6:,.2f} مليون).<br>'
             f'<b>التأجير:</b> {d["rent_billing_zero_b27"]["units"]} وحدةً '
             f"تفوتر وموازنتُها صفر "
             f'({d["rent_billing_zero_b27"]["annual"]/1e6:,.2f} مليون/سنة) · '
             f'{d["rent_vs_standard"]["no_standard"]} وحدةً بلا إيجارٍ معياريّ '
             f'({d["rent_vs_standard"]["no_standard_annual"]/1e6:,.2f} مليون/سنة).'
             "<br><b>التطوير:</b> جدولُ افتتاح المحطات بتاريخٍ مُلزم.</div>")

    with open(HTML, "w", encoding="utf-8") as f:
        f.write("".join(H))
    return HTML


def make_pdf():
    build_html()
    js = """const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch();
  const p = await b.newPage();
  await p.goto('file://' + process.argv[2], { waitUntil: 'networkidle' });
  await p.pdf({ path: process.argv[3], format: 'A4', printBackground: true });
  await b.close();
})();"""
    with open(JS, "w", encoding="utf-8") as f:
        f.write(js)
    subprocess.run(["node", JS, os.path.abspath(HTML), os.path.abspath(PDF)],
                   check=True)
    os.remove(HTML)          # وسيطٌ لا يُحفظ — المصدر هو هذا الملف
    os.remove(JS)
    print(f"✓ {PDF}")


if __name__ == "__main__":
    make_pdf() if "--html" not in sys.argv else print("✓", build_html())
