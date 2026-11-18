# -*- coding: utf-8 -*-
"""توقّعُ مبيعات المحطات الجديدة الـ٣٨ — شهراً بشهر، بتنسيق ورقة «Mo Bu»

   كلُّ رقمٍ في ورقة التوقّع معادلةٌ حيّةٌ تقرأ من ورقة «المُدخَلات»، فإن
   غيّرتَ معدّلَ النظير أو منحنى التدرّج تحرّك التوقّعُ كلُّه معه.

   والمُدخَلاتُ الثلاثة مقيسةٌ من بيانات درب وحدها — لا تقديرَ يُشبه القياس:

     ① معدّل النظير   — وسيطُ لتر/يوم لمحطاتِ المدينة نفسِها (وإلا المنطقة)
                        من ينا–سبت ٢٠٢٦ الفعليّ · ٩١ محطةً ناضجةً أساساً
     ② منحنى التدرّج  — مقيسٌ من ٢٢ افتتاحاً فعليّاً في ٢٠٢٦:
                        ٠٫٢٩ · ٠٫٦٨ · ٠٫٧٨ · ٠٫٩٠ · ٠٫٩٦ · ٠٫٩٨ · ١٫٠٠
     ③ مزيج المنتجات  — الحصصُ الفعليّةُ لمنطقة المحطة في ينا–سبت ٢٠٢٦

   وما لا نملكه مكتوبٌ صريحاً في ورقة «ما ينقص»: الحركةُ المارّةُ بكل موقع،
   وهي المُدخَلُ الذي يَحسم الفارق — ولا يُستعاض عنه بوسيطِ مدينة.

   يُشغَّل من جذر المستودع:  PYTHONPATH=data python3.13 data/build_new_stations.py
"""
import json

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.utils import get_column_letter as L

import brand as B
import worker_model as W

SRC = "outlets/data/new-stations-2027.json"
OUT = "docs/توقّع-المحطات-الجديدة-٢٠٢٧.xlsx"

MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
       "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
MONAR = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
         "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"]
PROD = [("91", "91"), ("95", "95"), ("diesel", "Diesel")]

# أعمدةُ ورقة التوقّع — ترتيبُ «Mo Bu» نفسُه
# A# B اسم C نموذج D نوع E ملاحظة F كود G سنة H منتج I..T شهور U إجمالي V متوسط W ل/يوم
C_MON = 9                 # أول عمودِ شهر
C_TOT, C_AVG, C_DAY, C_RATE, C_PCT = 21, 22, 23, 24, 25
NC = 25

# ورقةُ المُدخَلات — مواضعُ ثابتةٌ تُشار إليها من المعادلات
IN = "المُدخَلات"
R_RAMP = 6                # C6:I6  — منحنى التدرّج
R_DAYS = 8                # C8:N8  — أيام أشهر ٢٠٢٧
R_SEAS = 10               # C10:N10 — معامل الموسم (محايدٌ ١٫٠٠ افتراضاً)
R_REF = 11                # C11:N11 — المؤشّرُ المقيسُ مرجعاً (لا يُحسب منه)
R_ST = 15                 # أول صفٍّ في جدول المحطات

# المؤشّرُ الموسميُّ المقيس — ٨ محطاتٍ غربيةٍ لها سنتان فعليّتان كاملتان
# (وسيطُ ٢٠٢٥ و٢٠٢٦ لينا–سبت · وأكت–ديس من ٢٠٢٥ وحدها إذ لا فعليَّ لها في ٢٦)
SEAS_REF = [1.049, 1.003, 1.013, 1.012, 0.939, 0.926,
            1.022, 1.053, 0.982, 0.952, 0.920, 0.882]
# أعمدةُ جدول المحطات في ورقة المُدخَلات
IC = dict(no=1, code=2, name=3, reg=4, city=5, open=6, off=7, rate=8,
          src=9, n=10, m91=11, m95=12, mdsl=13)
MIXC = {"91": "m91", "95": "m95", "diesel": "mdsl"}

GOOD = PatternFill("solid", fgColor=B.T_GOOD)
BAD = PatternFill("solid", fgColor=B.T_BAD)
BAND = PatternFill("solid", fgColor=B.T_BAND)
NEU = PatternFill("solid", fgColor=B.T_NEUTRAL)
NUM = "#,##0;(#,##0);\"–\""
PCT = "+0.0%;-0.0%;\"–\""


def ramp_ref():
    return f"'{IN}'!$C${R_RAMP}:$I${R_RAMP}"


def days_ref(m):
    """خليّةُ أيامِ الشهر m (١..١٢)"""
    return f"'{IN}'!${L(2 + m)}${R_DAYS}"


def seas_ref(m):
    """معاملُ موسمِ الشهر m — مقسوماً على متوسّطِ الاثني عشر

       فالقسمةُ تجعله رافعةَ شكلٍ لا رافعةَ مستوى: يُعيد توزيعَ اللترات
       على الشهور ولا يغيّر الإجماليَّ السنويّ.
    """
    rng = f"'{IN}'!$C${R_SEAS}:$N${R_SEAS}"
    return f"'{IN}'!${L(2 + m)}${R_SEAS}/AVERAGE({rng})"


def fc_formula(ir, m, prod):
    """معادلةُ لترِ المنتجِ prod في الشهر m للمحطة التي مُدخَلاتها في الصف ir

       = معدّلُ النظير × التدرّج(شهرُ التشغيل) × أيامُ الشهر × حصّةُ المنتج

       وشهرُ التشغيل = m − «شهرُه»، حيث m شهرٌ ميلاديٌّ مرقَّمٌ من ١ و«شهرُه»
       موضعُ الافتتاح مرقَّمٌ من ٠ (٠ = يناير ٢٧ · −١ = ديسمبر ٢٦). فمحطةٌ
       تفتح في يناير ٢٧ شهرُها الأولُ هو m=١ ⇒ ١−٠ = ١. وما قبل الافتتاح صفر.
    """
    off = f"'{IN}'!$G${ir}"
    opm = f"{m}-{off}"
    mix = f"'{IN}'!${L(IC[MIXC[prod]])}${ir}"
    return (f"=IF({opm}<1,0,'{IN}'!$H${ir}"
            f"*INDEX({ramp_ref()},MIN(MAX({opm},1),7))"
            f"*{days_ref(m)}*{mix}*{seas_ref(m)})")


# ══════════════════════════════════════════════════════════════════
def sheet_inputs(wb, D):
    ws = wb.create_sheet(IN)
    W.setup(ws, [5, 9, 24, 11, 13, 12, 7, 11, 17, 5, 8, 8, 8], freeze="A15")
    W.title(ws, "المُدخَلات — أربعةُ مُدخَلاتٍ تُحرّك التوقّعَ كلَّه",
            "الأزرقُ يُغيَّر · والتوقّعُ في ورقة «FC 2027» معادلةٌ تقرأ من هنا، "
            "فما غيّرتَه هنا تحرّك هناك",
            13)
    W.rule(ws, 3, 13)

    # ① منحنى التدرّج
    c = ws.cell(4, 1, "① منحنى التدرّج — مقيسٌ من ٢٢ افتتاحاً فعليّاً في ٢٠٢٦")
    c.font = Font(name=B.FONT, size=10, bold=True, color=B.BGRAY)
    ws.cell(R_RAMP, 2, "شهرُ التشغيل ←").font = Font(name=B.FONT, size=9, color=B.INK2)
    for k, v in enumerate(D["ramp"]):
        hc = ws.cell(R_RAMP - 1, 3 + k, f"شهر {k + 1}")
        hc.font = Font(name=B.FONT, size=8.5, bold=True, color="FFFFFF")
        hc.fill = PatternFill("solid", fgColor=B.BGRAY)
        hc.alignment = W.CTR
        W.inp(ws, R_RAMP, 3 + k, "0.00", round(v, 3))
    ws.cell(R_RAMP, 11, "١٫٠٠ = مستوى الاستقرار · والشهرُ الأولُ ٠٫٢٩ لأنّه شهرٌ جزئيّ"
            ).font = Font(name=B.FONT, size=8.5, italic=True, color=B.INK2)

    # ② أيام الأشهر
    ws.cell(R_DAYS, 1, "② أيام ٢٠٢٧").font = Font(name=B.FONT, size=9.5, bold=True,
                                                   color=B.BGRAY)
    ws.cell(R_DAYS, 2, "فبراير ٢٨").font = Font(name=B.FONT, size=8.5, color=B.D_BAD)
    for m, d in enumerate(D["days"], 1):
        hc = ws.cell(R_DAYS - 1, 2 + m, MON[m - 1])
        hc.font = Font(name=B.FONT, size=8, bold=True, color="FFFFFF")
        hc.fill = PatternFill("solid", fgColor=B.BGRAY); hc.alignment = W.CTR
        W.inp(ws, R_DAYS, 2 + m, "0", d)

    # ③ معاملُ الموسم — رافعةٌ موصولةٌ ومتروكةٌ محايدةً عن قصد
    ws.cell(R_SEAS, 1, "③ معامل الموسم").font = Font(name=B.FONT, size=9.5,
                                                      bold=True, color=B.BGRAY)
    ws.cell(R_SEAS, 2, "محايدٌ ١٫٠٠").font = Font(name=B.FONT, size=8.5,
                                                   color=B.D_GOLD)
    for m in range(1, 13):
        W.inp(ws, R_SEAS, 2 + m, "0.000", 1.0)
    ws.cell(R_REF, 1, "مرجعاً — المقيس").font = Font(name=B.FONT, size=8.5,
                                                      color=B.INK2)
    ws.cell(R_REF, 2, "٨ محطاتٍ غربية").font = Font(name=B.FONT, size=8,
                                                     color=B.INK2)
    for m, v in enumerate(SEAS_REF, 1):
        cc = ws.cell(R_REF, 2 + m, v)
        cc.font = Font(name=B.FONT, size=8.5, italic=True, color=B.INK2)
        cc.border = W.BOX; cc.alignment = W.CTR; cc.number_format = "0.000"

    # ④ جدول المحطات
    r = 13
    ws.cell(r, 1, "④ لكل محطة: معدّلُ نظيرها ومزيجُها — والأساسُ مكتوبٌ بعدد النظراء")
    ws.cell(r, 1).font = Font(name=B.FONT, size=10, bold=True, color=B.BGRAY)
    W.header(ws, R_ST - 1, ["#", "الكود", "المحطة", "المنطقة", "المدينة", "الافتتاح",
                            "شهرُه*", "معدّل النظير\nل/يوم", "أساسُ المعدّل", "ن",
                            "٩١", "٩٥", "ديزل"], h=32)
    for i, s in enumerate(D["stations"], 1):
        rr = R_ST + i - 1
        W.calc(ws, rr, IC["no"], i, "0")
        W.calc(ws, rr, IC["code"], s["code"], "@")
        cc = ws.cell(rr, IC["name"], s["name"]); cc.font = W.BLACK
        cc.border = W.BOX; cc.alignment = W.RGT; cc.fill = NEU
        W.calc(ws, rr, IC["reg"], s["reg"], "@")
        W.calc(ws, rr, IC["city"], s["city"], "@")
        W.calc(ws, rr, IC["open"], s["open"], "@")
        W.calc(ws, rr, IC["off"], s["off"], "0")
        W.inp(ws, rr, IC["rate"], NUM, s["rate"])
        cc = ws.cell(rr, IC["src"], s["src"]); cc.font = Font(name=B.FONT, size=8.5,
                                                              color=B.INK2)
        cc.border = W.BOX; cc.alignment = W.RGT; cc.fill = NEU
        W.calc(ws, rr, IC["n"], s["n"], "0")
        for p in ("91", "95", "diesel"):
            W.inp(ws, rr, IC[MIXC[p]], "0.0%", s["mix"][p])
    last = R_ST + len(D["stations"]) - 1
    # نظيرٌ رقيقُ الأساس — تحذيرٌ بصريّ
    ws.conditional_formatting.add(
        f"{L(IC['n'])}{R_ST}:{L(IC['n'])}{last}",
        CellIsRule(operator="lessThan", formula=["5"], fill=BAD,
                   font=Font(name=B.FONT, size=10, bold=True, color=B.D_BAD)))
    ws.cell(last + 2, 1,
            "* «شهرُه» = موضعُ شهر الافتتاح نسبةً إلى يناير ٢٠٢٧ (٠ = يناير ٢٧ · "
            "−١ = ديسمبر ٢٦ · −٢ = نوفمبر ٢٦). ")
    ws.cell(last + 2, 1).font = Font(name=B.FONT, size=8.5, italic=True, color=B.INK2)
    ws.cell(last + 3, 1,
            "ن المظلَّلُ أحمر = نظراءُ أقلُّ من خمسٍ، فالوسيطُ هشٌّ ويحتاج حركةَ "
            "الموقع لا وسيطَ مدينة.")
    ws.cell(last + 3, 1).font = Font(name=B.FONT, size=8.5, italic=True, color=B.D_BAD)
    for j, t in enumerate((
            "معاملُ الموسم مُركَّبٌ ومتروكٌ على ١٫٠٠ عن قصد — لا لأنّ الموسمَ "
            "غيرُ موجود، بل لأنّه غيرُ مقيسٍ إلا على ثماني محطاتٍ غربية.",
            "والصفُّ المرجعيُّ أسفلَه هو ما قيس عليها (وسيطُ ٢٠٢٥ و٢٠٢٦): قاعُ "
            "مايو–يونيو ٠٫٩٣ وقمّةُ أغسطس ١٫٠٥ — وقاعُها شهرُ عيد الأضحى في "
            "السنتين (٦ يونيو ٢٠٢٥ · ٢٧ مايو ٢٠٢٦).",
            "وفي ٢٠٢٧ تتقدّم المواقيتُ: رمضان ٨ فبراير–٩ مارس · عيد الفطر ٩–١٠ "
            "مارس · عيد الأضحى ١٦–١٧ مايو. فمن نسخ الصفَّ المرجعيَّ كما هو فقد "
            "نسخ موضعَ ٢٠٢٦ لا ٢٠٢٧.",
            "والقسمةُ على متوسّطِ الاثني عشر تجعله رافعةَ شكلٍ لا مستوى: "
            "محطةٌ تعمل السنةَ كاملةً يتغيّر توزيعُها الشهريُّ ويثبت إجماليُّها، "
            "ومحطةٌ تفتح في أثناء السنة يتحرّك إجماليُّها بقدرِ موسمِ الأشهر "
            "التي عملتها وحدها.")):
        cc = ws.cell(last + 5 + j, 1, ("•  " if j else "⚠  ") + t)
        cc.font = Font(name=B.FONT, size=8.5,
                       italic=True, color=B.D_GOLD if j == 0 else B.INK2)
        ws.merge_cells(start_row=last + 5 + j, start_column=1,
                       end_row=last + 5 + j, end_column=13)
        cc.alignment = W.RGT
    last += 9

    # جداولُ مرجعيّة
    r = last + 5
    r = W.band(ws, r, 13, "وسيطُ النظراء كما قيس — ينا–سبت ٢٠٢٦ الفعليّ")
    W.header(ws, r, ["المستوى", "الاسم", "ل/يوم", "ن"], h=20)
    r += 1
    for k, (v, n) in sorted(D["peer_city"].items(), key=lambda x: -x[1][0]):
        W.calc(ws, r, 1, "مدينة", "@"); W.calc(ws, r, 2, k.title(), "@")
        W.calc(ws, r, 3, v, NUM); W.calc(ws, r, 4, n, "0"); r += 1
    for k, (v, n) in sorted(D["peer_reg"].items(), key=lambda x: -x[1][0]):
        W.calc(ws, r, 1, "منطقة", "@"); W.calc(ws, r, 2, k, "@")
        W.calc(ws, r, 3, v, NUM); W.calc(ws, r, 4, n, "0"); r += 1
    W.calc(ws, r, 1, "الشبكة", "@"); W.calc(ws, r, 2, "وسيط ٩١ ناضجة", "@")
    W.calc(ws, r, 3, D["peer_net"], NUM); W.calc(ws, r, 4, D["n_mature"], "0")
    r += 2

    r = W.band(ws, r, 13, "مزيجُ المنتجات الفعليُّ بالمنطقة — ينا–سبت ٢٠٢٦")
    W.header(ws, r, ["المنطقة", "٩١", "٩٥", "ديزل"], h=20)
    r += 1
    for rg, m in D["mix_reg"].items():
        W.calc(ws, r, 1, rg, "@")
        for j, p in enumerate(("91", "95", "diesel")):
            W.calc(ws, r, 2 + j, m[p], "0.0%")
        r += 1
    return ws


# ══════════════════════════════════════════════════════════════════
def sheet_fc(wb, D):
    ws = wb.create_sheet("FC 2027")
    W.setup(ws, [4, 26, 10, 12, 19, 8, 9, 8] + [11] * 12 + [12, 11, 10, 11, 10],
            freeze="I5")
    W.title(ws, "توقّعُ المحطات الجديدة الـ٣٨ — شهراً بشهر",
            "بترتيب «Mo Bu» · لكل محطةٍ كتلتان: موازنةُ الملف ثم التوقّعُ المقيس — "
            "والفرقُ أسفلَهما",
            NC)
    W.rule(ws, 3, NC)
    hrow = 4
    W.header(ws, hrow, ["#", "Station Name", "Model", "Type", "Note", "S.Code",
                        "Year", "Product"] + MON +
             ["Total", "Month Avrg", "Daily Avrg", "ل/يوم مستقرّ",
              "الفرق ÷\nالموازنة"], h=30)
    r = hrow + 1
    first = r
    bu_tot, fc_tot, vs_rows = [], [], []
    FCROW = {}
    for i, s in enumerate(D["stations"], 1):
        ir = R_ST + i - 1
        top = r
        # ── تقديرُ الملف لـ٢٠٢٦ (للسبع التي تفتح قبل ٢٠٢٧)
        if s["has26"]:
            cc = ws.cell(r, 7, "Est 2026")
            cc.font = Font(name=B.FONT, size=8.5, italic=True, color=B.INK2)
            cc.alignment = W.CTR; cc.border = W.BOX
            cc = ws.cell(r, 8, "Total")
            cc.font = Font(name=B.FONT, size=8.5, italic=True, color=B.INK2)
            cc.alignment = W.CTR; cc.border = W.BOX
            for m in range(1, 13):
                v = sum(s["a26"][p][m - 1] for p in ("91", "95", "diesel"))
                cc = ws.cell(r, C_MON - 1 + m, v or 0)
                cc.font = Font(name=B.FONT, size=9, italic=True, color=B.INK2)
                cc.border = W.BOX; cc.alignment = W.CTR; cc.number_format = NUM
            W.calc(ws, r, C_TOT, f"=SUM({L(C_MON)}{r}:{L(C_MON + 11)}{r})", NUM)
            r += 1
        # ── كتلتا الموازنة والتوقّع
        for lbl, kind in (("BU 2027", "bu"), ("FC 2027", "fc")):
            prow = {}
            for key, show in PROD:
                ws.cell(r, 7, lbl if key == "91" else None)
                ws.cell(r, 7).font = Font(name=B.FONT, size=9, bold=(key == "91"),
                                          color=B.BGRAY)
                ws.cell(r, 7).alignment = W.CTR
                cc = ws.cell(r, 8, show); cc.font = Font(name=B.FONT, size=9,
                                                         color=B.INK2)
                cc.alignment = W.CTR; cc.border = W.BOX
                for m in range(1, 13):
                    if kind == "bu":
                        v = s["bu"][key][m - 1]
                        cc = ws.cell(r, C_MON - 1 + m, v or 0)
                        cc.font = W.BLACK; cc.fill = NEU
                        cc.border = W.BOX; cc.alignment = W.CTR
                        cc.number_format = NUM
                    else:
                        W.calc(ws, r, C_MON - 1 + m, fc_formula(ir, m, key), NUM)
                W.calc(ws, r, C_TOT, f"=SUM({L(C_MON)}{r}:{L(C_MON + 11)}{r})", NUM)
                prow[key] = r
                r += 1
            # صفُّ الإجمالي
            ws.cell(r, 7, None)
            cc = ws.cell(r, 8, "Total"); cc.font = Font(name=B.FONT, size=9, bold=True,
                                                        color=B.BGRAY)
            cc.alignment = W.CTR; cc.border = W.BOX; cc.fill = BAND
            for m in range(1, 13):
                col = L(C_MON - 1 + m)
                f = "=" + "+".join(f"{col}{prow[k]}" for k, _ in PROD)
                cc = W.calc(ws, r, C_MON - 1 + m, f, NUM)
                cc.fill = BAND; cc.font = Font(name=B.FONT, size=9.5, bold=True,
                                               color=B.INK)
            for cnum, f, fmt in (
                    (C_TOT, f"=SUM({L(C_MON)}{r}:{L(C_MON + 11)}{r})", NUM),
                    (C_AVG, f"={L(C_TOT)}{r}/12", NUM),
                    (C_DAY, f"={L(C_TOT)}{r}/365", NUM),
                    (C_RATE, f"=MAX({L(C_MON)}{r}:{L(C_MON + 11)}{r})/31", NUM)):
                cc = W.calc(ws, r, cnum, f, fmt)
                cc.fill = BAND; cc.font = Font(name=B.FONT, size=9.5, bold=True,
                                               color=B.INK)
            (bu_tot if kind == "bu" else fc_tot).append(r)
            if kind == "fc":
                FCROW[s["code"]] = r
            r += 1
        # ── الفرق
        bt, ft = bu_tot[-1], fc_tot[-1]
        cc = ws.cell(r, 8, "الفرق"); cc.font = Font(name=B.FONT, size=9, bold=True,
                                                     color=B.D_BAD)
        cc.alignment = W.CTR; cc.border = W.BOX
        ws.cell(r, 7, "FC − BU").font = Font(name=B.FONT, size=8.5, color=B.INK2)
        ws.cell(r, 7).alignment = W.CTR
        for m in range(1, 13):
            col = L(C_MON - 1 + m)
            W.calc(ws, r, C_MON - 1 + m, f"={col}{ft}-{col}{bt}", NUM)
        W.calc(ws, r, C_TOT, f"={L(C_TOT)}{ft}-{L(C_TOT)}{bt}", NUM)
        W.calc(ws, r, C_AVG, f"={L(C_TOT)}{r}/12", NUM)
        W.calc(ws, r, C_PCT, f"=IF({L(C_TOT)}{bt}=0,\"\","
                             f"{L(C_TOT)}{ft}/{L(C_TOT)}{bt}-1)", PCT)
        vs_rows.append(r)
        # ── عمودُ الهويّة: يُكتب مرّةً ويُدمج على الكتلة
        for cnum, val in ((1, i), (2, s["name"]), (3, s["model"]),
                          (4, "New Station"), (5, s["src"]), (6, s["code"])):
            ws.merge_cells(start_row=top, start_column=cnum,
                           end_row=r, end_column=cnum)
            cc = ws.cell(top, cnum, val)
            cc.font = (Font(name=B.FONT, size=10, bold=True, color=B.INK) if cnum == 2
                       else Font(name=B.FONT, size=8.5, color=B.INK2))
            cc.alignment = W.CTR if cnum != 2 else W.RGT
            cc.border = W.BOX
            if cnum == 2:
                cc.fill = NEU
        for j in range(1, NC + 1):
            ws.cell(r, j).border = W.BOX
        r += 1

    # ── كتلةُ الإجمالي العام
    r = W.band(ws, r, NC, "Grand Total — 38 stations")
    for lbl, rows in (("BU 2027", bu_tot), ("FC 2027", fc_tot)):
        cc = ws.cell(r, 8, lbl); cc.font = Font(name=B.FONT, size=10, bold=True,
                                                color="FFFFFF")
        cc.fill = PatternFill("solid", fgColor=B.BGRAY); cc.alignment = W.CTR
        cc.border = W.BOX
        ws.cell(r, 2, "المجموع").font = Font(name=B.FONT, size=10, bold=True,
                                              color=B.BGRAY)
        for m in range(1, 13):
            col = L(C_MON - 1 + m)
            f = "=" + "+".join(f"{col}{x}" for x in rows)
            cc = W.calc(ws, r, C_MON - 1 + m, f, NUM)
            cc.font = Font(name=B.FONT, size=9.5, bold=True, color=B.INK)
        for cnum, f in ((C_TOT, f"=SUM({L(C_MON)}{r}:{L(C_MON + 11)}{r})"),
                        (C_AVG, f"={L(C_TOT)}{r}/12"),
                        (C_DAY, f"={L(C_TOT)}{r}/365")):
            cc = W.calc(ws, r, cnum, f, NUM)
            cc.font = Font(name=B.FONT, size=9.5, bold=True, color=B.INK)
        r += 1
    gbu, gfc = r - 2, r - 1
    cc = ws.cell(r, 8, "الفرق"); cc.font = Font(name=B.FONT, size=10, bold=True,
                                                 color=B.D_BAD)
    cc.alignment = W.CTR; cc.border = W.BOX
    for m in range(1, 13):
        col = L(C_MON - 1 + m)
        W.calc(ws, r, C_MON - 1 + m, f"={col}{gfc}-{col}{gbu}", NUM)
    W.calc(ws, r, C_TOT, f"={L(C_TOT)}{gfc}-{L(C_TOT)}{gbu}", NUM)
    W.calc(ws, r, C_AVG, f"={L(C_TOT)}{r}/12", NUM)
    W.calc(ws, r, C_PCT, f"={L(C_TOT)}{gfc}/{L(C_TOT)}{gbu}-1", PCT)
    ws.cell(r, 2, "التوقّعُ المقيس ÷ موازنة الملف").font = Font(
        name=B.FONT, size=9.5, bold=True, color=B.D_BAD)
    last = r

    # تنسيقٌ شرطيّ: الفروقُ السالبة حمراء والموجبة خضراء
    rng = f"{L(C_MON)}{first}:{L(C_TOT)}{last}"
    ws.conditional_formatting.add(rng, CellIsRule(
        operator="lessThan", formula=["0"], fill=BAD,
        font=Font(name=B.FONT, size=9, color=B.D_BAD)))
    for op, fl, col in (("lessThan", BAD, B.D_BAD),
                        ("greaterThan", GOOD, B.D_GOOD)):
        ws.conditional_formatting.add(
            f"{L(C_PCT)}{first}:{L(C_PCT)}{last}",
            CellIsRule(operator=op, formula=["0"], fill=fl,
                       font=Font(name=B.FONT, size=9, bold=True, color=col)))
    return ws, gbu, gfc, FCROW


# ══════════════════════════════════════════════════════════════════
def sheet_gap(wb, D, FCROW):
    ws = wb.create_sheet("الفجوة")
    W.setup(ws, [4, 26, 11, 12, 8, 11, 11, 12, 12, 11, 10, 22], freeze="A5")
    W.title(ws, "الفجوة — أين يتركّز الفارقُ بين الموازنة والتوقّع",
            "مرتَّبةً بالفارق · والعمودُ الأخيرُ يقول ما يَحسم كلَّ سطر",
            12)
    W.rule(ws, 3, 12)
    W.header(ws, 4, ["#", "المحطة", "الكود", "الافتتاح", "أشهرُ\nالتشغيل",
                     "موازنة ٢٧\nم لتر", "التوقّع\nم لتر", "الفارق\nم لتر",
                     "التوقّع ÷\nالموازنة", "ل/يوم\nنظير", "ل/يوم\nتُوازِنه",
                     "ما يَحسمه"], h=34)
    rows = []
    for i, s in enumerate(D["stations"], 1):
        bu = sum(sum(s["bu"][p]) for p in ("91", "95", "diesel"))
        act = [m for m in range(12) if m >= s["off"]]
        mo = len(act)
        rt_bu = max(sum(s["bu"][p][m] for p in s["bu"]) for m in range(12)) / 31
        rows.append((i, s, bu, mo, rt_bu))
    # التوقّعُ لا يُكرَّر هنا — يُقرأ من صفِّ الإجمالي في ورقة «FC 2027»
    r = 5
    order = sorted(rows, key=lambda x: -x[2])
    for n, (i, s, bu, mo, rt_bu) in enumerate(order, 1):
        f = FCROW[s["code"]]
        W.calc(ws, r, 1, n, "0")
        cc = ws.cell(r, 2, s["name"]); cc.font = Font(name=B.FONT, size=10,
                                                      color=B.INK)
        cc.border = W.BOX; cc.alignment = W.RGT; cc.fill = NEU
        W.calc(ws, r, 3, s["code"], "@")
        W.calc(ws, r, 4, s["open"], "@")
        W.calc(ws, r, 5, mo, "0")
        W.calc(ws, r, 6, bu / 1e6, "0.00")
        W.calc(ws, r, 7, f"='FC 2027'!U{f}/1000000", "0.00")
        W.calc(ws, r, 8, f"={L(7)}{r}-{L(6)}{r}", "+0.00;-0.00")
        W.calc(ws, r, 9, f"=IF({L(6)}{r}=0,\"\",{L(7)}{r}/{L(6)}{r})", "0%")
        W.calc(ws, r, 10, f"='{IN}'!$H${R_ST + i - 1}", NUM)
        W.calc(ws, r, 11, rt_bu, NUM)
        note = ("الحركةُ المارّةُ بالموقع — وسيطُ المدينة لا يَحسمها"
                if s["n"] >= 5 else f"نظراءُ {s['n']} فقط — الأساسُ هشّ")
        if rt_bu > 90000:
            note = "موازنةٌ بمعدّلٍ يفوق ٩٠ ألف ل/يوم — تحتاج عدّادَ حركةٍ لا قالباً"
        cc = ws.cell(r, 12, note)
        cc.font = Font(name=B.FONT, size=8.5, color=B.D_BAD if rt_bu > 90000
                       else B.INK2)
        cc.border = W.BOX; cc.alignment = W.RGT
        r += 1
    tot = r
    cc = ws.cell(r, 2, "المجموع — ٣٨ محطة")
    cc.font = Font(name=B.FONT, size=10, bold=True, color="FFFFFF")
    cc.fill = PatternFill("solid", fgColor=B.BGRAY); cc.alignment = W.RGT
    cc.border = W.BOX
    for cnum, f, fmt in ((6, f"=SUM(F5:F{r - 1})", "0.0"),
                         (7, f"=SUM(G5:G{r - 1})", "0.0"),
                         (8, f"=G{r}-F{r}", "+0.0;-0.0"),
                         (9, f"=G{r}/F{r}", "0%")):
        cc = W.calc(ws, r, cnum, f, fmt)
        cc.fill = BAND; cc.font = Font(name=B.FONT, size=10, bold=True, color=B.INK)
    ws.conditional_formatting.add(f"H5:H{tot}", CellIsRule(
        operator="lessThan", formula=["0"], fill=BAD,
        font=Font(name=B.FONT, size=9.5, bold=True, color=B.D_BAD)))
    ws.conditional_formatting.add(f"H5:H{tot}", CellIsRule(
        operator="greaterThan", formula=["0"], fill=GOOD,
        font=Font(name=B.FONT, size=9.5, bold=True, color=B.D_GOOD)))
    ws.conditional_formatting.add(f"I5:I{tot - 1}", ColorScaleRule(
        start_type="num", start_value=0, start_color=B.T_BAD,
        mid_type="num", mid_value=1, mid_color="FFFFFF",
        end_type="num", end_value=2, end_color=B.T_GOOD))
    ws.conditional_formatting.add(f"K5:K{tot - 1}", CellIsRule(
        operator="greaterThan", formula=["90000"], fill=BAD,
        font=Font(name=B.FONT, size=9.5, bold=True, color=B.D_BAD)))
    return ws


# ══════════════════════════════════════════════════════════════════
def sheet_gaps_known(wb, D):
    ws = wb.create_sheet("ما ينقص")
    W.setup(ws, [4, 34, 30, 26, 22], freeze="A5")
    W.title(ws, "ما ينقص — ولا يُملأ بتقديرٍ يُشبه القياس",
            "هذه الورقةُ جزءٌ من النموذج: المُدخَلُ الناقصُ مكتوبٌ بمالكه، "
            "لا بفرضٍ يبدو رقماً",
            5)
    W.rule(ws, 3, 5)
    W.header(ws, 4, ["#", "المُدخَل الناقص", "لماذا يَحسم", "من يملك سدَّه",
                     "أثرُه على هذه الورقة"], h=32)
    ROWS = [
        ("الحركةُ المارّةُ بكل موقعٍ جديد (زيارة/يوم)",
         "الوحدةُ تبيع ما يمرّ بها — ووسيطُ المدينة لا يعرف طريقَ المحطة",
         "هندسةُ الشبكة",
         "يستبدل العمودَ «معدّل النظير» كلَّه بقياسٍ موقعيّ"),
        ("أشهرٌ فعليّةٌ كاملةٌ لـ٢٠٢٥ لغير المنطقة الغربية",
         "الثماني الوحيدةُ التي لها سنتان كلُّها غربية — فلا موسمَ مقيسٌ للجنوبية "
         "والوسطى والشرقية",
         "الإدارة المالية",
         "يسمح بإسقاطِ منحنى موسميٍّ على الشهور بدل توزيعٍ مستوٍ"),
        ("تاريخُ افتتاحٍ للـ٤٢ محطةً خارج هذا الشيت",
         "تحمل ٤٣٢ مليون لتر (١٧٪ من موازنة ٢٧) وليست في «New Stations 2027»",
         "إدارةُ المشاريع",
         "تُضاف ككتلةٍ ثانيةٍ بالطريقة نفسها"),
        ("سببُ انكسار المعيصم وعرفات الشرائع في فبراير ٢٠٢٦",
         "−٦٧٪ و−٦٦٪ في ليلةٍ واحدة — وهو ما يفسّر فارقَ ٣٠ نقطةً في طلب القائم",
         "إدارةُ التشغيل",
         "يحدّد هل معدّلُ نظراءِ مكة صالحٌ أساساً للجديدة"),
        ("سعرُ الشراء وهامشُ اللتر لكل منتج",
         "الورقةُ باللترات فقط — والقرارُ يُتخذ بالريال",
         "الإدارةُ المالية",
         "يحوّل اللترات إلى هامشٍ ويرتّب الأولويات بالريال لا باللتر"),
    ]
    r = 5
    for i, (k, why, own, eff) in enumerate(ROWS, 1):
        W.calc(ws, r, 1, i, "0")
        for cnum, v, col in ((2, k, B.INK), (3, why, B.INK2), (4, own, B.D_GOLD),
                             (5, eff, B.INK2)):
            cc = ws.cell(r, cnum, v)
            cc.font = Font(name=B.FONT, size=9, bold=(cnum == 2), color=col)
            cc.border = W.BOX
            cc.alignment = W.WRAP
        ws.row_dimensions[r].height = 34
        r += 1
    r += 1
    W.bullets(ws, r, 5, [(
        "كيف تُقرأ هذه الورقة",
        ["التوقّعُ هنا أفضلُ قراءةٍ تحتملها بياناتنا — لا قياسُ موقع",
         "وثلاثةُ مُدخَلاته مقيسة: معدّلُ النظير ومنحنى التدرّج ومزيجُ المنطقة",
         "وما ليس مقيساً مكتوبٌ أعلاه بمالكه — ولم يُملأ برقمٍ يبدو قياساً"],
        B.BGRAY)])
    return ws


def build():
    D = json.load(open(SRC, encoding="utf-8"))
    wb = Workbook(); wb.remove(wb.active)
    sheet_inputs(wb, D)
    _, gbu, gfc, FCROW = sheet_fc(wb, D)
    sheet_gap(wb, D, FCROW)
    sheet_gaps_known(wb, D)
    wb._sheets = [wb["FC 2027"], wb["الفجوة"], wb[IN], wb["ما ينقص"]]
    wb.save(OUT)
    print(f"✓ {OUT}")
    print(f"  محطات {len(D['stations'])} · منحنى التدرّج "
          f"{' · '.join(f'{x:.2f}' for x in D['ramp'])}")


if __name__ == "__main__":
    build()
