# -*- coding: utf-8 -*-
"""انحراف كل محطة شهراً بشهر — ٢٠٢٥ مقابل ٢٠٢٦

   السؤال ليس «كم تغيّرت المحطة» بل «كم تغيّرت بما لا يفسّره غيرُها».
   ولذلك لا يُقرأ رقمٌ خام هنا: كل انحرافٍ يمرّ بثمانية محدِّدات قبل أن
   يُسمّى انحرافاً — وورقة «المحددات» تسمّيها وتقول أين يحملها الجدول.

   وأهمّها ثلاثة أثبتها هذا الملف نفسه:
     · النظير يُحسب من المحطات **الناضجة** وبدون المحطة المقيسة — وإلا
       رفع نموُّ محطةٍ جديدة سقفَ الجميع فبدت الناضجةُ كلُّها متراجعة.
     · أكتوبر–ديسمبر ٢٠٢٦ تقديرٌ لا فعل — فلا انحراف فيها.
     · عمود VS% في الورقة الأصلية مقسومٌ على السنة الجديدة — فـ«‎−٨١٪»
       تغيُّرُه الحقيقي ‎−٤٥٪.

   يُشغَّل من جذر المستودع:  PYTHONPATH=data python3 data/build_deviation.py
"""
import json

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.formatting.rule import FormulaRule, ColorScaleRule
from openpyxl.utils import get_column_letter as L

import brand as B
import worker_model as W

SRC = "outlets/data/mobu.json"
OUT = "docs/انحراف-المحطات-شهر-بشهر.xlsx"

M = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
     "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"]
PCT = "0%"
PCT1 = "+0.0%;-0.0%;0%"
NUM = "#,##0"

# ── المحددات الثمانية: (المحدِّد, ما يفسده إن أُهمل, كيف يُعالَج, أين في الجدول)
DET = [
 ("أيام الشهر",
  "فبراير ٢٨ ومارس ٣١ — فرقٌ ١١٪ بلا أي تغيّرٍ في الأداء",
  "القسمة على أيام الشهر قبل أي مقارنة",
  "صفُّ «أيام الشهر» في ورقة البيانات · وكل الكتل لتر/يوم"),
 ("الشهر فعليٌّ أم مُقدَّر",
  "أكتوبر–ديسمبر ٢٠٢٦ تقديرٌ متساوٍ حرفياً — الانحراف فيها صناعةُ ملف",
  "تُستبعد من القياس ويُكتفى بيناير–سبتمبر",
  "الجدول تسعةُ أشهرٍ لا اثنا عشر · والمستبعَد مكتوبٌ في البيانات"),
 ("اكتمال الأساس",
  "العمرة بدأت نوفمبر ٢٥ — نسبتها ‎+٩٢٢٪ وهي ليست نمواً",
  "لا تُقاس محطةٌ ينقص أساسُها، وتُستبعد من النظير أيضاً",
  "عمود «الحالة»: ناضجة · أساسٌ ناقص · بلا حجم"),
 ("النظير — ومن يدخله",
  "لو دخل النظيرَ نموُّ محطةٍ جديدة ارتفع سقفُ الجميع فبدت الناضجة متراجعة",
  "النظير = الناضجة وحدها، وبدون المحطة المقيسة",
  "كتلة «النظير» · وصفّا «مجموع الناضجة» في البيانات"),
 ("تركّز الحجم",
  "محطةٌ كبيرةٌ تقيس نفسها بنفسها إن بقيت داخل النظير",
  "طرحُ المحطة من بسط النظير ومقامه",
  "صيغة النظير: (مجموع الناضجة ناقص هذه المحطة)"),
 ("مزيج المنتج ٩١ · ٩٥ · ديزل",
  "تحوّلٌ بين المنتجات يغيّر القيمة والهامش دون أن يغيّر الحجم",
  "الحجم يُقاس هنا، والمزيج يُفحص على حدة قبل أي حكمٍ على الإيراد",
  "ورقة البيانات تحمل المنتجات · والجدول حجمٌ فقط"),
 ("التقويم الهجري",
  "رمضان ٢٠٢٥ في مارس و٢٠٢٦ من ١٨ فبراير — فبراير بفبراير يخلط موسمين",
  "الأشهر التي يتحرّك فيها الموسم تُعلَّم ولا يُبنى عليها حكمٌ منفرد",
  "كتلة الانحراف الصافي · وفبراير ومارس تُقرآن معاً"),
 ("حدّ الدلالة",
  "نسبةٌ ثابتةٌ ٥٪ تُنذر في محطةٍ هادئة وتصمت في محطةٍ متذبذبة",
  "الحدّ من تذبذب المحطة نفسها: الانحراف المعياري لانحرافها الصافي",
  "عمود σ · وعمود الحكم · والمضاعف خليةٌ تُعدَّل"),
]

NOTES = [
 "الناضجة التسع ‎−١٢٫٢٪ في يناير–سبتمبر، والشبكة ‎+١٤٫١٪ — والفرق كلُّه العمرة.",
 "المعيصم وعرفات الشرائع وحدهما ‎−١٤٫٥ مليون لتر = ٨٧٪ من تراجع الناضجة.",
 "وكلتاهما انكسرت في الشهر نفسه: فبراير ٢٦ ‎−٦٠٪ و‎−٥١٪ عن يناير.",
 "والسبع الباقية ‎−١٫٩٪ — أي داخل تذبذبها، لا مشكلة بيعٍ فيها.",
 "محطتان تنكسران في شهرٍ واحدٍ بالقدر نفسه ليستا مشكلة مبيعات: يُصنَّف السبب أولاً.",
]


def build():
    d = json.load(open(SRC, encoding="utf-8"))
    ACT = d["act_months_26"]
    D25, D26 = d["days25"], d["days26"]
    S = d["stations"]
    live = [s for s in S if s["live"]]
    mat = [s for s in live if s["mature"]]
    other = [s for s in live if not s["mature"]]
    mat.sort(key=lambda s: -sum(s["a26"][:ACT]))
    wb = Workbook()
    wb.remove(wb.active)

    # ═════════════════ ① المحددات
    ws = wb.create_sheet("المحددات")
    NC = 8
    W.setup(ws, [4, 22, 44, 40, 40, 10, 10, 10], freeze="A5")
    W.title(ws, "المحددات — ثمانيةٌ قبل أن يُسمّى الرقم انحرافاً",
            "الرقم الخام لا يُقرأ: فرقٌ بين شهرين قد يكون أيامَ الشهر لا أداءَ المحطة",
            NC)
    W.rule(ws, 3, NC)
    W.header(ws, 4, ["#", "المحدِّد", "ما يفسده إن أُهمل", "كيف يُعالَج",
                     "أين يحمله الجدول", "", "", ""])
    ws.merge_cells(start_row=4, start_column=5, end_row=4, end_column=NC)
    for x in range(5, NC + 1):
        ws.cell(4, x).fill = W.FH
        ws.cell(4, x).border = W.BOX
    for i, (a, b, c, e) in enumerate(DET):
        r = 5 + i
        ws.cell(r, 1, i + 1).alignment = W.CTR
        ws.cell(r, 2, a).font = W.BOLD
        ws.cell(r, 2).alignment = W.RGT
        ws.cell(r, 3, b).font = Font(name=W.F, size=10, color=B.D_BAD)
        ws.cell(r, 3).alignment = W.WRAP
        ws.cell(r, 4, c).font = Font(name=W.F, size=10, color=B.D_GOOD)
        ws.cell(r, 4).alignment = W.WRAP
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=NC)
        ws.cell(r, 5, e).font = W.SMALL
        ws.cell(r, 5).alignment = W.WRAP
        for x in range(1, NC + 1):
            ws.cell(r, x).border = W.BOX
        ws.row_dimensions[r].height = 34
    r = 5 + len(DET) + 1
    W.bullets(ws, r, NC, [("ما يقوله الجدول بعد تطبيقها", NOTES, B.D_BAD)])

    # ═════════════════ ② البيانات
    bs = wb.create_sheet("البيانات")
    BC = 18
    W.setup(bs, [4, 24, 9, 12, 13, 9] + [11] * 12, freeze="A6")
    W.title(bs, "البيانات — لتراً بالشهر",
            f"المصدر: {d['source']} › ورقة Mo Bu · والمطابقة مع صفِّ الإجمالي فيها صفر",
            BC)
    W.rule(bs, 3, BC)
    hdr = ["#", "المحطة", "الكود", "النموذج", "الحالة", "السنة"] + M
    W.header(bs, 4, hdr)
    # صفُّ أيام الشهر
    DROW = 5
    bs.cell(DROW, 2, "أيام الشهر").font = W.BOLD
    bs.cell(DROW, 2).alignment = W.RGT
    for k in range(12):
        c = bs.cell(DROW, 7 + k, D26[k])
        c.number_format = NUM
        c.alignment = W.CTR
        c.font = W.BOLD
        c.fill = W.FS
    for x in range(1, BC + 1):
        bs.cell(DROW, x).border = W.BOX
        if x < 7:
            bs.cell(DROW, x).fill = W.FS

    rowmap = {}
    r = DROW + 1
    for i, s in enumerate(live, 1):
        state = ("ناضجة" if s["mature"] else
                 ("أساسٌ ناقص" if sum(s["a26"]) > 0 else "بلا حجم"))
        for tag, yr, dd in (("a25", "٢٠٢٥", D25), ("a26", "٢٠٢٦", D26)):
            bs.cell(r, 1, i if tag == "a25" else None).alignment = W.CTR
            bs.cell(r, 2, s["name"].split(" - ")[0] if tag == "a25" else None)
            bs.cell(r, 2).font = W.BOLD
            bs.cell(r, 2).alignment = W.RGT
            bs.cell(r, 3, s["code"] if tag == "a25" else None).alignment = W.CTR
            bs.cell(r, 4, s["model"] if tag == "a25" else None).alignment = W.CTR
            c = bs.cell(r, 5, state if tag == "a25" else None)
            c.alignment = W.CTR
            if tag == "a25":
                c.fill = {"ناضجة": W.OK, "أساسٌ ناقص": W.WARN}.get(state, W.BAD)
            bs.cell(r, 6, yr).alignment = W.CTR
            bs.cell(r, 6).font = W.BOLD
            for k in range(12):
                cc = bs.cell(r, 7 + k, s[tag][k])
                cc.number_format = NUM
                cc.alignment = W.CTR
                if tag == "a26" and k >= ACT:
                    cc.fill = W.WARN
                    cc.font = Font(name=W.F, size=10, color=B.D_GOLD)
            for x in range(1, BC + 1):
                bs.cell(r, x).border = W.BOX
            rowmap[(s["code"], tag)] = r
            r += 1
    # مجموع الناضجة
    TM = {}
    for tag, yr in (("a25", "٢٠٢٥"), ("a26", "٢٠٢٦")):
        bs.cell(r, 2, f"مجموع الناضجة {yr}").font = W.BOLD
        bs.cell(r, 2).alignment = W.RGT
        for k in range(12):
            refs = "+".join(f"{L(7+k)}{rowmap[(s['code'], tag)]}" for s in mat)
            cc = bs.cell(r, 7 + k, f"={refs}")
            cc.number_format = NUM
            cc.alignment = W.CTR
            cc.font = W.BOLD
            cc.fill = W.FS
        for x in range(1, BC + 1):
            bs.cell(r, x).border = W.BOX
            if x < 7:
                bs.cell(r, x).fill = W.FS
        TM[tag] = r
        r += 1
    r += 1
    W.bullets(bs, r, BC, [("ملاحظات المصدر", [
        "أكتوبر–ديسمبر ٢٠٢٦ (المظلَّل) تقديرٌ في الملف الأصلي لا فعل",
        "٢٠٢٥ من «2025 Results.xlsx» · و٢٠٢٦ و٢٠٢٧ صيغٌ حيّة على دفتر GC",
        "عمود VS% في الورقة الأصلية مقسومٌ على السنة الجديدة لا على الأساس — "
        "فـ«‎−٨١٪» تغيُّره الحقيقي ‎−٤٥٪",
    ], B.D_GOLD)])

    # ═════════════════ ③ الانحراف
    ds = wb.create_sheet("الانحراف شهر بشهر")
    NC2 = 14
    W.setup(ds, [4, 22] + [9] * 9 + [11, 8, 14], freeze="C7")
    W.title(ds, "الانحراف شهراً بشهر — بعد المحددات",
            "لتر/يوم · يناير–سبتمبر ٢٠٢٦ مقابل ٢٠٢٥ · والنظير الناضجة بدون المحطة",
            NC2)
    W.rule(ds, 3, NC2)
    ds.cell(4, 2, "مضاعف حدّ الدلالة (× σ)").font = W.BOLD
    ds.cell(4, 2).alignment = W.RGT
    W.inp(ds, 4, 3, "0.0", 1.0)
    ds.merge_cells(start_row=4, start_column=4, end_row=4, end_column=NC2)
    c = ds.cell(4, 4, "حدّ الدلالة = σ ÷ √٩ × المضاعف — خطأُ متوسطِ التسعة "
                      "أشهر لا تذبذبُ شهرٍ واحد")
    c.font = W.SMALL
    c.alignment = W.RGT
    MULT = "$C$4"

    blocks = [("① لتر/يوم ٢٠٢٥", NUM), ("② لتر/يوم ٢٠٢٦", NUM),
              ("③ الانحراف الخام ٪", PCT), ("④ النظير ٪ — الناضجة بدون المحطة", PCT),
              ("⑤ الانحراف الصافي — نقطة مئوية", PCT1)]
    start = {}
    r = 6
    for bi, (title, fmt) in enumerate(blocks):
        r = W.band(ds, r, NC2, title)
        cols = ["#", "المحطة"] + [m[:3] for m in M[:ACT]]
        cols += (["السنة", "σ", "الحكم"] if bi == 4 else ["السنة", "", ""])
        W.header(ds, r, cols)
        hr = r
        r += 1
        start[bi] = r
        for i, s in enumerate(mat):
            rr = r + i
            ds.cell(rr, 1, i + 1).alignment = W.CTR
            ds.cell(rr, 2, s["name"].split(" - ")[0]).font = W.BOLD
            ds.cell(rr, 2).alignment = W.RGT
            r25, r26 = rowmap[(s["code"], "a25")], rowmap[(s["code"], "a26")]
            for k in range(ACT):
                col = L(3 + k)
                dcol = L(7 + k)
                if bi == 0:
                    f = f"='البيانات'!{dcol}{r25}/'البيانات'!{dcol}${DROW}"
                elif bi == 1:
                    f = f"='البيانات'!{dcol}{r26}/'البيانات'!{dcol}${DROW}"
                elif bi == 2:
                    f = (f'=IF({col}{start[0]+i}=0,"—",'
                         f"{col}{start[1]+i}/{col}{start[0]+i}-1)")
                elif bi == 3:
                    f = (f"=('البيانات'!{dcol}${TM['a26']}-'البيانات'!{dcol}{r26})"
                         f"/('البيانات'!{dcol}${TM['a25']}-'البيانات'!{dcol}{r25})-1")
                else:
                    f = (f'=IF(ISNUMBER({col}{start[2]+i}),'
                         f'{col}{start[2]+i}-{col}{start[3]+i},"—")')
                cc = W.calc(ds, rr, 3 + k, f, fmt)
                if bi == 4:
                    cc.font = W.BOLD
            # عمود السنة
            yc = L(3 + ACT)
            if bi == 0:
                f = (f"=SUM('البيانات'!G{r25}:{L(6+ACT)}{r25})"
                     f"/SUM('البيانات'!G${DROW}:{L(6+ACT)}${DROW})")
            elif bi == 1:
                f = (f"=SUM('البيانات'!G{r26}:{L(6+ACT)}{r26})"
                     f"/SUM('البيانات'!G${DROW}:{L(6+ACT)}${DROW})")
            elif bi == 2:
                f = f"={yc}{start[1]+i}/{yc}{start[0]+i}-1"
            elif bi == 3:
                f = (f"=(SUM('البيانات'!G${TM['a26']}:{L(6+ACT)}${TM['a26']})"
                     f"-SUM('البيانات'!G{r26}:{L(6+ACT)}{r26}))"
                     f"/(SUM('البيانات'!G${TM['a25']}:{L(6+ACT)}${TM['a25']})"
                     f"-SUM('البيانات'!G{r25}:{L(6+ACT)}{r25}))-1")
            else:
                f = f"={yc}{start[2]+i}-{yc}{start[3]+i}"
            cc = W.calc(ds, rr, 3 + ACT, f, fmt)
            cc.font = W.BOLD
            if bi == 4:
                sc = L(3 + ACT + 1)
                W.calc(ds, rr, 3 + ACT + 1,
                       f"=STDEVP(C{rr}:{L(2+ACT)}{rr})", PCT)
                hk = (f'=IF(ABS({yc}{rr})<={sc}{rr}/SQRT({ACT})*{MULT},'
                      f'"داخل التذبذب",'
                      f'IF({yc}{rr}>0,"فوق النظير","تحت النظير"))')
                cc = W.calc(ds, rr, 3 + ACT + 2, hk, "General")
                cc.font = W.BOLD
            else:
                for x in (3 + ACT + 1, 3 + ACT + 2):
                    ds.cell(rr, x).border = W.BOX
            for x in range(1, NC2 + 1):
                ds.cell(rr, x).border = W.BOX
            ds.row_dimensions[rr].height = 18
        r += len(mat) + 1
        if bi == 4:
            rng = f"C{start[4]}:{L(3+ACT)}{start[4]+len(mat)-1}"
            ds.conditional_formatting.add(rng, ColorScaleRule(
                start_type="num", start_value=-0.3, start_color=B.T_BAD,
                mid_type="num", mid_value=0, mid_color="FFFFFF",
                end_type="num", end_value=0.3, end_color=B.T_GOOD))
            hrng = f"{L(3+ACT+2)}{start[4]}:{L(3+ACT+2)}{start[4]+len(mat)-1}"
            ds.conditional_formatting.add(hrng, FormulaRule(
                formula=[f'{L(3+ACT+2)}{start[4]}="تحت النظير"'], fill=W.BAD,
                font=Font(name=W.F, size=10, bold=True, color=B.D_BAD)))
            ds.conditional_formatting.add(hrng, FormulaRule(
                formula=[f'{L(3+ACT+2)}{start[4]}="فوق النظير"'], fill=W.OK,
                font=Font(name=W.F, size=10, bold=True, color=B.D_GOOD)))

    # المستبعَدات
    r = W.band(ds, r, NC2, "⑥ خارج القياس — ولماذا")
    W.header(ds, r, ["#", "المحطة", "السبب"] + [""] * (NC2 - 3))
    ds.merge_cells(start_row=r, start_column=4, end_row=r, end_column=NC2)
    for x in range(4, NC2 + 1):
        ds.cell(r, x).fill = W.FH
        ds.cell(r, x).border = W.BOX
    excl = [(s["name"].split(" - ")[0],
             ("بلا حجمٍ في السنتين — تفتح " + (s["note"] or "لاحقاً"))
             if sum(s["a26"]) == 0 else
             (f"بدأت {M[s['first25']]} ٢٠٢٥ — الأساس ناقص"
              if s["first25"] is not None else "بلا أساسٍ في ٢٠٢٥"))
            for s in other]
    excl += [(s["name"].split(" - ")[0],
              "بلا حجمٍ في السنتين — تفتح " + (s["note"] or "لاحقاً"))
             for s in S if not s["live"]]
    for i, (nm, why) in enumerate(excl):
        rr = r + 1 + i
        ds.cell(rr, 1, i + 1).alignment = W.CTR
        ds.cell(rr, 2, nm).font = W.BOLD
        ds.cell(rr, 2).alignment = W.RGT
        ds.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=NC2)
        ds.cell(rr, 3, why).alignment = W.RGT
        ds.cell(rr, 3).font = Font(name=W.F, size=10, color=B.D_GOLD)
        for x in range(1, NC2 + 1):
            ds.cell(rr, x).border = W.BOX
        ds.row_dimensions[rr].height = 18
    r += len(excl) + 2
    W.bullets(ds, r, NC2, [("القراءة", NOTES, B.D_BAD)])

    wb.save(OUT)
    print(f"✓ {OUT} · {len(wb.sheetnames)} أوراق")
    print(f"  ناضجة: {len(mat)} · خارج القياس: {len(excl)} · "
          f"أشهرٌ فعلية: {ACT}")


if __name__ == "__main__":
    build()
