# -*- coding: utf-8 -*-
"""موازنة ٢٠٢٧ بالكتل الثلاث — الرقم يتحرّك مع التاريخ

   ثلاث كتلٍ لا تُخلط، لأن لكلٍّ مالكاً ومخاطرةً مختلفة:

     ① القائم          — معدّل المحطة الفعليّ × نموٍّ يُقرَّر   · مدير المبيعات
     ② أثر السنة الكاملة — محطةٌ فُتحت في ٢٦ تُشغَّل ١٢ شهراً   · حسابٌ آليّ
     ③ الافتتاحات       — معدّل نظيرة × تدرّج × أشهر التشغيل  · التطوير

   والمُدخَلات أربعة فقط: نموُّ القائم · منحنى التدرّج · تاريخ افتتاح كل
   محطة · ومعدّلها المفترض. فمن غيّر تاريخاً تحرّكت الموازنة كلُّها، ولا
   يُدافَع عن رقمٍ كُتب مرّة.

   والقاعدة معدّل **يوليو–سبتمبر ٢٠٢٦ الفعليّ** لا ديسمبر: ديسمبر نفسه
   مُقدَّرٌ في الملف الأصلي، والبناء على تقديرٍ يضاعف خطأه.

   يُشغَّل من جذر المستودع:  PYTHONPATH=data python3 data/build_budget27.py
"""
import json
import statistics as st

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.utils import get_column_letter as L

import brand as B
import worker_model as W

SRC = "outlets/data/mobu.json"
OUT = "docs/موازنة-٢٠٢٧-بالكتل-الثلاث.xlsx"
M = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
     "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"]
D27 = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
BASE = (6, 9)                      # نافذة الأساس: يوليو–سبتمبر ٢٠٢٦
RAMP = [0.40, 0.60, 0.75, 0.85, 0.95, 1.00]
NUM = "#,##0"
MLN = "#,##0.0"
PCT = "0%"
MONTH_OF = {"Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
            "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12}


def open_month(note):
    """«Jun.2027» ← ٦ · «2028» أو «Dec.2026» ← خارج السنة"""
    t = (note or "").strip()
    if not t:
        return None
    if "2027" not in t:
        return 0 if "2026" in t else 99       # ٠ = تعمل من يناير · ٩٩ = بعد ٢٧
    for k, v in MONTH_OF.items():
        if t.lower().startswith(k.lower()):
            return v
    return 99


def build():
    d = json.load(open(SRC, encoding="utf-8"))
    S = d["stations"]
    D26 = d["days26"]
    b0, b1 = BASE
    days_base = sum(D26[b0:b1])
    for s in S:
        s["rate"] = sum(s["a26"][b0:b1]) / days_base
        s["file27"] = sum(s["b27"])
        s["om"] = open_month(s["note"])
    live = [s for s in S if s["rate"] > 0]
    rayth = [s for s in S if "الريث" in s["name"]]
    newly = [s for s in S if s["rate"] <= 0 and s not in rayth]
    med = st.median([s["rate"] for s in live])
    for s in rayth + newly:
        s["assume"] = med                     # الافتراض الأوّليّ: وسيط الشبكة
        if s["om"] is None:
            s["om"] = 99
    for s in rayth:
        s["om"] = 1                           # فُتحت ديسمبر ٢٦ — تعمل السنة كاملة
    order = [("① القائم", live), ("② أثر السنة الكاملة", rayth),
             ("③ الافتتاحات", newly)]

    wb = Workbook()
    wb.remove(wb.active)

    # ═══════════════ المُدخَلات
    ws = wb.create_sheet("المُدخَلات")
    NC = 12
    W.setup(ws, [4, 24, 10, 13, 15, 15, 13, 13, 14, 10, 10, 10], freeze="A12")
    W.title(ws, "المُدخَلات — أربعةٌ تُحرّك الموازنة كلَّها",
            "الملوّن يُعدَّل · والباقي محسوبٌ من الفعليّ", NC)
    W.rule(ws, 3, NC)
    r = W.band(ws, 4, NC, "① مُدخَلاتٌ عامة")
    GEN = [("نموُّ المحطة القائمة", 0.0, "0%",
            "قرارٌ لا افتراض — صفرٌ يعني تثبيت المعدّل الفعليّ"),
           ("تأخيرُ الافتتاحات (أشهر)", 0, "0",
            "يُزاح كل افتتاحٍ بهذا العدد — لاختبار الخطة"),
           ("نافذة الأساس", "يوليو–سبتمبر ٢٠٢٦", "General",
            "آخر ثلاثة أشهرٍ فعلية — وديسمبر مُقدَّرٌ فلا يصلح قاعدة")]
    for i, (k, v, fmt, why) in enumerate(GEN):
        rr = r + i
        ws.cell(rr, 2, k).font = W.BOLD
        ws.cell(rr, 2).alignment = W.RGT
        c = W.inp(ws, rr, 3, fmt, v)
        ws.merge_cells(start_row=rr, start_column=4, end_row=rr, end_column=NC)
        cc = ws.cell(rr, 4, why)
        cc.font = W.SMALL
        cc.alignment = W.RGT
        for x in range(2, NC + 1):
            ws.cell(rr, x).border = W.BOX
    GROW, DELAY = "$C$5", "$C$6"
    r += len(GEN) + 1

    r = W.band(ws, r, NC, "② منحنى التدرّج — نسبة المعدّل الكامل في أشهر التشغيل")
    hdr = ["", "الشهر"] + [f"{i+1}" for i in range(6)] + ["٧ فما بعد"]
    for j, t in enumerate(hdr, 1):
        if t:
            c = ws.cell(r, j, t)
            c.font = W.HEAD
            c.fill = W.FH
            c.alignment = W.CTR
            c.border = W.BOX
    ws.cell(r, 2, "الشهر").font = W.HEAD
    r += 1
    ws.cell(r, 2, "نسبة التشغيل").font = W.BOLD
    ws.cell(r, 2).alignment = W.RGT
    ws.cell(r, 2).border = W.BOX
    for i in range(6):
        W.inp(ws, r, 3 + i, "0%", RAMP[i])
    W.inp(ws, r, 9, "0%", 1.0)
    RAMPROW = r
    r += 1
    r = W.note_or(ws, r, NC) if False else r
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=NC)
    c = ws.cell(r, 2, "لا محطةَ تبلغ معدّلها في شهرها الأول. وهذا المنحنى "
                      "افتراضٌ حتى يُستخرج من آخر خمس افتتاحاتٍ فعلية — "
                      "ودليلُنا اليوم محطةٌ واحدة.")
    c.font = Font(name=W.F, size=9, color=B.D_GOLD)
    c.fill = W.WARN
    c.alignment = W.WRAP
    for x in range(3, NC + 1):
        ws.cell(r, x).fill = W.WARN
        ws.cell(r, x).border = W.BOX
    ws.cell(r, 2).border = W.BOX
    ws.row_dimensions[r].height = 26
    r += 2

    r = W.band(ws, r, NC, "③ المحطات")
    W.header(ws, r, ["#", "المحطة", "الكود", "النموذج", "الكتلة",
                     "معدّل الأساس", "المعدّل المفترض", "شهر الافتتاح",
                     "معدّل الملف الضمنيّ", "", "", ""])
    r += 1
    SROW = {}
    i = 0
    for blk, group in order:
        for s in group:
            i += 1
            rr = r + i - 1
            ws.cell(rr, 1, i).alignment = W.CTR
            ws.cell(rr, 2, s["name"].split(" - ")[0]).font = W.BOLD
            ws.cell(rr, 2).alignment = W.RGT
            ws.cell(rr, 3, s["code"]).alignment = W.CTR
            ws.cell(rr, 4, s["model"]).alignment = W.CTR
            ws.cell(rr, 5, blk).alignment = W.CTR
            c = W.calc(ws, rr, 6, s["rate"] if s["rate"] > 0 else 0, NUM)
            c.value = round(s["rate"]) if s["rate"] > 0 else None
            if blk == "① القائم":
                ws.cell(rr, 7, None).border = W.BOX
                ws.cell(rr, 8, None).border = W.BOX
            else:
                W.inp(ws, rr, 7, NUM, round(s["assume"]))
                W.inp(ws, rr, 8, "0", s["om"] if s["om"] != 99 else None)
            days = sum(D27[k] for k in range(12) if s["b27"][k] > 0) or 365
            c = W.calc(ws, rr, 9, round(s["file27"] / days) if s["file27"] else 0,
                       NUM)
            c.value = round(s["file27"] / days) if s["file27"] else None
            for x in range(1, NC + 1):
                ws.cell(rr, x).border = W.BOX
            ws.row_dimensions[rr].height = 18
            SROW[s["code"] or s["name"]] = rr
    r += i + 1
    W.bullets(ws, r, NC, [
        ("كيف تُقرأ", [
            "«معدّل الأساس» محسوبٌ من الفعليّ ولا يُعدَّل — وهو ما تبيعه المحطة اليوم",
            "«المعدّل المفترض» للمحطة التي لا تاريخ لها — وأوّلُه وسيط الشبكة",
            "«شهر الافتتاح» ١–١٢ · والفراغ يعني لا تفتح في ٢٠٢٧",
            "«معدّل الملف الضمنيّ» ما تفترضه موازنة الملف — للمقارنة لا للاستعمال",
        ], B.BGRAY),
        ("فرقٌ يستحق قراراً", [
            f"وسيط الشبكة العاملة {med:,.0f} لتر/يوم · "
            "وموازنة الملف تفترض للمحطة الجديدة ١١٠–١١٧ ألفاً = نحو ٢٫٤×",
            "ولا تبلغها المحطة من شهرها الأول في الملف — بلا تدرّج",
        ], B.D_BAD),
    ])

    # ═══════════════ الموازنة
    bs = wb.create_sheet("الموازنة")
    BC = 19
    W.setup(bs, [4, 22, 15, 11] + [10] * 12 + [12], freeze="E6")
    W.title(bs, "موازنة ٢٠٢٧ — بالكتل الثلاث",
            "مليون لتر · كلُّ خليةٍ صيغةٌ في المُدخَلات، فمن غيّر تاريخاً تحرّك الرقم",
            BC)
    W.rule(bs, 3, BC)
    W.header(bs, 4, ["#", "المحطة", "الكتلة", "لتر/يوم"] + [m[:3] for m in M]
             + ["السنة"])
    r = 5
    blkrows = {}
    i = 0
    for blk, group in order:
        first = r
        for s in group:
            i += 1
            sr = SROW[s["code"] or s["name"]]
            bs.cell(r, 1, i).alignment = W.CTR
            bs.cell(r, 2, s["name"].split(" - ")[0]).font = W.BOLD
            bs.cell(r, 2).alignment = W.RGT
            bs.cell(r, 3, blk).alignment = W.CTR
            if blk == "① القائم":
                rate = f"'المُدخَلات'!F{sr}*(1+'المُدخَلات'!{GROW})"
                W.calc(bs, r, 4, f"={rate}", NUM)
                for k in range(12):
                    W.calc(bs, r, 5 + k, f"=$D{r}*{D27[k]}/1000000", MLN)
            else:
                rate = f"'المُدخَلات'!G{sr}"
                W.calc(bs, r, 4, f"={rate}", NUM)
                for k in range(12):
                    idx = (f"{k+1}-'المُدخَلات'!H{sr}-'المُدخَلات'!{DELAY}+1")
                    f = (f'=IF(\'المُدخَلات\'!H{sr}="",0,'
                         f"IF({idx}<1,0,"
                         f"$D{r}*{D27[k]}/1000000*"
                         f"INDEX('المُدخَلات'!$C${RAMPROW}:$I${RAMPROW},"
                         f"MIN({idx},7))))")
                    W.calc(bs, r, 5 + k, f, MLN)
            W.calc(bs, r, 17 + 0, f"=SUM(E{r}:P{r})", MLN)
            for x in range(1, BC + 1):
                bs.cell(r, x).border = W.BOX
            bs.row_dimensions[r].height = 18
            r += 1
        bs.cell(r, 2, f"مجموع {blk}").font = W.BOLD
        bs.cell(r, 2).alignment = W.RGT
        for k in range(13):
            c = W.calc(bs, r, 5 + k, f"=SUM({L(5+k)}{first}:{L(5+k)}{r-1})", MLN)
            c.font = W.BOLD
            c.fill = W.FS
        for x in range(1, BC + 1):
            bs.cell(r, x).border = W.BOX
            if x < 5:
                bs.cell(r, x).fill = W.FS
        blkrows[blk] = r
        r += 2
    # الإجمالي
    bs.cell(r, 2, "إجمالي الموازنة").font = Font(name=W.F, size=11, bold=True,
                                                 color=B.ORANGE)
    bs.cell(r, 2).alignment = W.RGT
    for k in range(13):
        f = "+".join(f"{L(5+k)}{blkrows[b]}" for b, _ in order)
        c = W.calc(bs, r, 5 + k, f"={f}", MLN)
        c.font = Font(name=W.F, size=11, bold=True, color=B.ORANGE)
        c.fill = W.SECT
    for x in range(1, BC + 1):
        bs.cell(r, x).border = W.BOX
        if x < 5:
            bs.cell(r, x).fill = W.SECT
    TOT = r
    r += 2

    # المقارنة
    r = W.band(bs, r, BC, "المقارنة — موازنتنا مقابل موازنة الملف")
    W.header(bs, r, ["", "البند", "", "", "مليون لتر"] + [""] * 13 + [""])
    file27 = sum(s["file27"] for s in S) / 1e6
    floor = sum(s["rate"] for s in live) * 365 / 1e6
    r += 1
    for lbl, val, note in (
            ("موازنة هذا الملف", f"=Q{TOT}", "بمُدخَلاتها الحالية"),
            ("موازنة الملف الأصلي", round(file27, 1), "BU 2027 في ورقة Mo Bu"),
            ("الأرضية — بلا افتتاحٍ ولا نمو", round(floor, 1),
             "معدّل يوليو–سبتمبر × ٣٦٥"),
            ("الفرق عن الملف الأصلي", f"=E{r}-E{r+1}", "موجبٌ = خطّتنا أعلى")):
        bs.cell(r, 2, lbl).font = W.BOLD
        bs.cell(r, 2).alignment = W.RGT
        c = W.calc(bs, r, 5, val, MLN)
        c.font = W.BOLD
        bs.merge_cells(start_row=r, start_column=6, end_row=r, end_column=BC)
        cc = bs.cell(r, 6, note)
        cc.font = W.SMALL
        cc.alignment = W.RGT
        for x in range(1, BC + 1):
            bs.cell(r, x).border = W.BOX
        r += 1
    r += 1

    # الحساسية — صفٌّ لكل سيناريو، والخلية مجموع المحطات في ذلك الشهر
    r = W.band(bs, r, BC, "الحساسية — ماذا لو تأخّرت الافتتاحات")
    W.header(bs, r, ["", "التأخير", "", ""] + [m[:3] for m in M]
             + ["الكتلتان ② و③"])
    r += 1
    newrows = [(s, SROW[s["code"] or s["name"]])
               for blk, group in order if blk != "① القائم" for s in group]
    first_sens = r
    for dmonths in (0, 1, 3, 6):
        bs.cell(r, 2, f"تأخير {dmonths} شهر" if dmonths else "بلا تأخير")
        bs.cell(r, 2).font = W.BOLD
        bs.cell(r, 2).alignment = W.RGT
        for k in range(12):
            terms = []
            for s, sr in newrows:
                idx = f"{k+1}-'المُدخَلات'!H{sr}-{dmonths}+1"
                terms.append(
                    f'IF(\'المُدخَلات\'!H{sr}="",0,IF({idx}<1,0,'
                    f"'المُدخَلات'!G{sr}*{D27[k]}/1000000*"
                    f"INDEX('المُدخَلات'!$C${RAMPROW}:$I${RAMPROW},"
                    f"MIN({idx},7))))")
            W.calc(bs, r, 5 + k, "=" + "+".join(terms), MLN)
        c = W.calc(bs, r, BC, f"=SUM(E{r}:P{r})", MLN)
        c.font = W.BOLD
        c.fill = W.FS
        for x in range(1, BC + 1):
            bs.cell(r, x).border = W.BOX
        r += 1
    bs.cell(r, 2, "الفاقد عن «بلا تأخير»").font = W.BOLD
    bs.cell(r, 2).alignment = W.RGT
    for i, dm in enumerate((0, 1, 3, 6)):
        pass
    c = W.calc(bs, r, BC, f"={L(BC)}{first_sens+3}-{L(BC)}{first_sens}", MLN)
    c.font = Font(name=W.F, size=10, bold=True, color=B.D_BAD)
    c.fill = W.BAD
    bs.merge_cells(start_row=r, start_column=5, end_row=r, end_column=BC - 1)
    cc = bs.cell(r, 5, "تأخيرٌ ستةُ أشهرٍ وسطياً — كم لتراً يسقط")
    cc.font = W.SMALL
    cc.alignment = W.RGT
    for x in range(1, BC + 1):
        bs.cell(r, x).border = W.BOX
    r += 2
    W.bullets(bs, r, BC, [("ما يُقرأ قبل التوقيع", [
        "كم تساوي الموازنة لو لم تُفتح محطةٌ واحدة؟ — سطرُ «الأرضية»",
        "كم من النمو بيعٌ وكم إنشاء؟ — كتلة ① مقابل ② و③",
        "الحسينية تعمل بمعدّلٍ فعليٍّ وموازنةُ الملف لها صفر — تُراجَع",
        "الريث موازنتها في الملف مبنيّةٌ على ديسمبر ٢٦ وهو شهرٌ مُقدَّر",
    ], B.D_BAD)])

    wb.save(OUT)
    print(f"✓ {OUT}")
    print(f"  قائم {len(live)} · سنةٌ كاملة {len(rayth)} · افتتاحات {len(newly)}")
    print(f"  وسيط المعدّل {med:,.0f} لتر/يوم · الأرضية {floor:,.1f}م · "
          f"موازنة الملف {file27:,.1f}م")


if __name__ == "__main__":
    build()
