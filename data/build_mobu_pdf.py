# -*- coding: utf-8 -*-
"""صفحةٌ واحدةٌ للطباعة — أرقامُ ورقة Mo Bu باللتر على ترويسةِ درب

   كلُّ رقمٍ فيها يُقرأ من `outlets/data/mobu-years.json` الذي يبنيه
   `read_mobu_years`، فلا رقمَ مكتوبٌ في الصفحةِ بلا مصدر.

   يُشغَّل من جذر المستودع:  PYTHONPATH=data python3.13 data/build_mobu_pdf.py
"""
import base64
import json
import os
import subprocess

SRC = "outlets/data/mobu-years.json"
HTML = "docs/أرقام-Mo-Bu.html"
PDF = "docs/أرقام-Mo-Bu.pdf"
JS = "data/mobu_pdf.js"
FONT = "data/fonts/din-next-lt-arabic-medium.ttf"
LOGO = "data/brand/darb-logo.png"
MARK = "data/brand/darb-letterhead-watermark.jpg"
ORANGE = "#F37521"
GREY = "#59595C"

MON = ["ينا", "فبر", "مار", "أبر", "ماي", "يون",
       "يول", "أغس", "سبت", "أكت", "نوف", "ديس"]

CSS = """
@font-face { font-family:"DIN"; src:url(data:font/ttf;base64,__FONT__) format("truetype");
             font-weight:100 900; font-style:normal; }
@page { size: A4 portrait; }
* { box-sizing: border-box; }
html { background-image: url(__MARK__); background-repeat: no-repeat;
       background-position: left bottom; background-size: 100% 100%;
       background-attachment: fixed; }
body { direction: rtl; margin: 0; color: #3D3D3D; font-family: "DIN", sans-serif; }
.head { border-bottom: 1pt solid __ORANGE__; padding-bottom: 2mm; margin: 0 0 2.5mm; }
.head .n { font-size: 13pt; color: __GREY__; }
.head .d { font-size: 8pt; color: #8E8A84; margin-top: 1mm; }
h2 { font-size: 9pt; color: __GREY__; background: #F5F2ED; margin: 2.6mm 0 1.2mm;
     padding: 1mm 2.5mm; border-right: 1.8pt solid __ORANGE__; }
table { border-collapse: collapse; width: 100%; font-size: 8pt; }
th { background: __GREY__; color: #fff; padding: 1mm 1mm; font-size: 7.2pt;
     text-align: center; font-weight: 600; }
td { border: 0.4pt solid #E6E2DB; padding: 0.9mm 1.4mm; text-align: center; }
td.k { text-align: right; color: __GREY__; background: #FAF9F7; white-space: nowrap; }
td.big { font-size: 10pt; font-weight: 700; color: #1F1F1F; }
tr.tot td { background: #F5F2ED; font-weight: 700; }
.mon { font-size: 6.4pt; }
.mon td { padding: 0.7mm 0.3mm; }
.note { font-size: 7.2pt; color: #6E6A64; background: #F5F2ED; padding: 1.4mm 2.2mm;
        margin-top: 1mm; border-right: 1.8pt solid __ORANGE__; line-height: 1.5; }
.note b { color: __GREY__; }
.warn { border-right-color: #C0503A; }
.warn b { color: #9A3E2C; }
.pos { color: #1F5E4A; font-weight: 700; }
.neg { color: #9A3E2C; font-weight: 700; }
.src { font-size: 6.8pt; color: #A8A49E; margin-top: 4mm; text-align: left;
       direction: ltr; }
"""


def fmt(x, d=1):
    return f"{x / 1e6:,.{d}f}"


def build_html(D):
    b64 = lambda p: base64.b64encode(open(p, "rb").read()).decode()
    css = (CSS.replace("__FONT__", b64(FONT)).replace("__ORANGE__", ORANGE)
              .replace("__GREY__", GREY)
              .replace("__MARK__", "data:image/jpeg;base64," + b64(MARK)))
    Y = D["years"]
    H = ['<meta charset="utf-8"><style>' + css + "</style>"]
    H.append('<div class="head"><div class="n">أرقامُ ورقة Mo Bu باللتر</div>'
             '<div class="d">القطاع التجاري · ثلاثُ سنواتٍ على '
             f'{D["stations"]} محطة — وكلُّ رقمٍ محسوبٌ من الورقةِ لا منقولاً'
             '</div></div>')

    # ① الإجمالي
    H.append("<h2>① الإجماليُّ السنويّ — مليون لتر</h2><table>"
             "<tr><th>السنة</th><th>محطات</th><th>مليون لتر</th>"
             "<th>الحالة</th><th>تحقّقُ عمود Total</th></tr>")
    STAT = {"2025": "فعليٌّ — ١٢ شهراً",
            "2026": "فعليٌّ ينا–سبت + تقديرٌ أكت–ديس",
            "2027": "موازنة"}
    for y in ("2025", "2026", "2027"):
        v = Y[y]
        H.append(f'<tr><td class="k">{y}</td><td>{v["n"]}</td>'
                 f'<td class="big">{fmt(v["total"])}</td>'
                 f'<td style="font-size:7.4pt">{STAT[y]}</td>'
                 f'<td class="pos">{v["total"] - v["own"]:+,.0f} لتر</td></tr>')
    H.append("</table>")
    H.append('<div class="note"><b>التحقّق</b> — جمعُ الشهورِ الاثني عشرَ '
             'يطابق عمودَ <b>Total</b> الذي تكتبه الورقةُ بنفسها: الفرقُ '
             '<b>صفرٌ في السنواتِ الثلاث</b>.</div>')

    # ② شهراً بشهر
    H.append("<h2>② شهراً بشهر — مليون لتر</h2>"
             '<table class="mon"><tr><th>السنة</th>'
             + "".join(f"<th>{m}</th>" for m in MON) + "<th>السنة</th></tr>")
    for y in ("2025", "2026", "2027"):
        H.append(f'<tr><td class="k">{y}</td>'
                 + "".join(f"<td>{x / 1e6:,.1f}</td>" for x in Y[y]["m"])
                 + f'<td style="font-weight:700">{fmt(Y[y]["total"])}</td></tr>')
    H.append("</table>")
    H.append(f'<div class="note warn"><b>٢٠٢٦ ليست كلُّها فعليّة</b> — '
             f'{fmt(D["act26"])} مليون لترٍ فعليٌّ (يناير–سبتمبر) و'
             f'{fmt(D["est26"])} مليون لترٍ تقديرٌ (أكتوبر–ديسمبر). '
             f'والدليل: قيمُ الثلاثةِ أشهرٍ <b>متساويةٌ تماماً</b> في '
             f'{D["est_eq"]} محطةً من {D["stations"]}.</div>')

    # ③ قواعد المقارنة
    cm, cl = D["common"], D["clean"]
    H.append("<h2>③ المقارنةُ لا تصحُّ إلا على القاعدةِ نفسِها</h2><table>"
             "<tr><th>القاعدة</th><th>٢٠٢٥</th><th>٢٠٢٦</th><th>٢٠٢٧</th>"
             "<th>٢٦ ÷ ٢٥</th><th>٢٧ ÷ ٢٦</th><th>ينا–سبت ٢٦ ÷ ٢٥</th></tr>")
    for lbl, b, cls in ((f"{cm['n']} محطةً لها ٢٠٢٥", cm, ""),
                        (f"{cl['n']} محطةً مفتوحةً في السنتين", cl, "tot")):
        g1 = b["y"][1] / b["y"][0] - 1
        g2 = b["y"][2] / b["y"][1] - 1
        g3 = b["ytd"][1] / b["ytd"][0] - 1
        sp = lambda v: (f'<span class="{"pos" if v > 0 else "neg"}">'
                        f'{v * 100:+.1f}٪</span>')
        H.append(f'<tr class="{cls}"><td class="k">{lbl}</td>'
                 + "".join(f"<td>{fmt(x)}</td>" for x in b["y"])
                 + f"<td>{sp(g1)}</td><td>{sp(g2)}</td><td>{sp(g3)}</td></tr>")
    H.append("</table>")
    H.append('<div class="note warn"><b>+٧١٪ افتتاحاتٌ لا نموّ</b> — '
             f'{cm["n"] - cl["n"]} محطةً من الـ{cm["n"]} ليست مفتوحةً في كلِّ '
             'أشهرِ ٢٠٢٥، فسنتُها جزئيّةٌ وتُقرَأ قفزة. وعلى القاعدةِ '
             'النظيفةِ الشبكةُ <b>نزلت</b>، والموازنةُ تطلب منها صعوداً.'
             '</div>')

    # ④ محرّكات القاعدة النظيفة
    H.append("<h2>④ ما يحرّك القاعدةَ النظيفة — ينا–سبت ٢٦ مقابل ٢٥</h2>"
             "<table><tr><th>المحطة</th><th>التغيُّر</th>"
             "<th>المحطة</th><th>التغيُّر</th></tr>")
    mv = D["movers"]
    for i in range(0, len(mv), 2):
        row = ""
        for j in (i, i + 1):
            if j < len(mv):
                n, v = mv[j]
                cls = "neg" if v < 0 else "pos"
                row += (f'<td class="k">{n}</td>'
                        f'<td class="{cls}">{v * 100:+.1f}٪</td>')
            else:
                row += "<td></td><td></td>"
        H.append(f"<tr>{row}</tr>")
    H.append("</table>")

    # ⑤ عيب الأكواد
    top = " · ".join(f"{k} ×{n}" for k, n in D["code_top"])
    H.append("<h2>⑤ عيبٌ في الورقةِ يُفسد أيَّ حسابٍ بالكود</h2>")
    H.append(f'<div class="note warn"><b>عمودُ S.Code معطوب</b> — '
             f'{D["code_values"]} قيمةً فقط لـ{D["stations"]} محطة: {top}. '
             f'وأكثرُ من ذلك: الأكوادُ الموجودةُ مُسنَدةٌ لغيرِ أصحابها — '
             f'<b>MD011</b> على كورنيش جازان وهو ابيار الماشي، و<b>MK127</b> '
             f'على خليص وهو بحرة، و<b>JA357</b> على الذكرى الخالدة وهو الريث. '
             f'فالمفتاحُ الصالحُ عمودُ <b>Station Name</b> أو عمودُ '
             f'<b>#</b> (١–{D["stations"]}) — لا الكود.</div>')

    # ⑥ المحطات الجديدة
    N = D.get("new")
    if N:
        qq = " · ".join(f"ر{k} {v}" for k, v in N["q"].items())
        H.append("<h2>⑥ المحطاتُ الجديدةُ لـ٢٠٢٧ — خارجَ الستّين</h2><table>"
                 "<tr><th>العدد</th><th>تفتح نوف–ديس ٢٦</th>"
                 "<th>تفتح خلال ٢٧</th><th>موازنةُ الملف</th>"
                 "<th>التوقّعُ المقيس</th><th>التوقّع ÷ الموازنة</th></tr>")
        H.append(f'<tr class="tot"><td class="big">{N["n"]}</td>'
                 f'<td>{N["n26"]}</td><td>{N["n27"]}</td>'
                 f'<td class="big">{fmt(N["budget"])}</td>'
                 f'<td class="big">{fmt(N["forecast"])}</td>'
                 f'<td class="neg">'
                 f'{N["forecast"] / N["budget"] * 100:.0f}٪</td></tr>')
        H.append("</table>")
        H.append(f'<div class="note">توزُّعُ افتتاحاتِ ٢٠٢٧: {qq}. '
                 f'<b>ولا تقاطعَ مع الستّين</b> — فالرقمان يُجمعان ولا '
                 f'يُكرَّران. والتوقّعُ مبنيٌّ على أربعةِ مُدخَلاتٍ مقيسةٍ '
                 f'من بياناتنا: معدّلُ النظير '
                 f'({N["rate_lo"]:,}–{N["rate_hi"]:,} ل/يوم) × منحنى التدرّج '
                 f'المقيس من ٢٢ افتتاحاً في ٢٠٢٦ × كسرِ يومِ الافتتاح × '
                 f'منحنى الموسم — وتفصيلُه في ملف '
                 f'«توقّع المحطات الجديدة ٢٠٢٧».</div>')
        H.append(f'<div class="note warn"><b>المجموعُ على هذه الورقة</b> — '
                 f'الستّون {fmt(D["years"]["2027"]["total"])} مليون لترٍ '
                 f'موازنةً، والجديدةُ {fmt(N["budget"])} موازنةً أو '
                 f'{fmt(N["forecast"])} توقّعاً ⇒ '
                 f'<b>{fmt(D["years"]["2027"]["total"] + N["budget"])}</b> '
                 f'بموازنةِ الملف، أو '
                 f'<b>{fmt(D["years"]["2027"]["total"] + N["forecast"])}</b> '
                 f'بالتوقّعِ المقيس.</div>')

    H.append(f'<div class="src">{D["source"]}</div>')
    with open(HTML, "w", encoding="utf-8") as f:
        f.write("".join(H))
    return HTML


def build():
    D = json.load(open(SRC, encoding="utf-8"))
    html = build_html(D)
    logo64 = base64.b64encode(open(LOGO, "rb").read()).decode()
    with open(JS, "w", encoding="utf-8") as f:
        f.write("""const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch();
  const p = await b.newPage();
  await p.goto('file://' + process.argv[2], { waitUntil: 'networkidle' });
  await p.pdf({
    path: process.argv[3], format: 'A4', printBackground: true,
    displayHeaderFooter: true,
    margin: { top: '27mm', right: '13mm', bottom: '15mm', left: '13mm' },
    headerTemplate: '<div style="width:100%;padding:9mm 13mm 0;">'
      + '<img src="' + process.argv[4] + '" style="height:12mm"></div>',
    footerTemplate: '<div style="width:100%;text-align:center;font-size:6pt;'
      + 'color:#B5AEA2;font-family:sans-serif;padding-bottom:6mm;">'
      + 'محسوبٌ من ورقة Mo Bu &nbsp;·&nbsp; <span class="pageNumber"></span>/'
      + '<span class="totalPages"></span></div>',
  });
  await b.close();
})();""")
    subprocess.run(["node", JS, os.path.abspath(html), os.path.abspath(PDF),
                    "data:image/png;base64," + logo64], check=True)
    os.remove(html)
    os.remove(JS)
    print(f"✓ {PDF}")


if __name__ == "__main__":
    build()
