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
from openpyxl.worksheet.datavalidation import DataValidation

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
R_RAMP = 6                # C6:I6   — منحنى التدرّج · K6 كسرُ عيّنةِ القياس
C_FRAC = 11               # K6      — كسرُ شهرِ الافتتاح في عيّنةِ القياس
R_DAYS = 8                # C8:N8   — أيام أشهر ٢٠٢٧
R_PROF = 11               # B11:O14 — أربعةُ منحنياتِ موسمٍ · O المتوسّط
N_PROF = 4
C_PAVG = 15               # O       — متوسّطُ المنحنى (للتطبيع)
R_ANC = 18                # C18:G18 — رافعةُ المستأجرين: ميلٌ وثابتٌ وإشغالٌ وسقف
C_SLOPE, C_INTER, C_OCC, C_CAP = 3, 4, 5, 6
R_ST = 23                 # أول صفٍّ في جدول المحطات
BASES = ["نظيرُ الجوار", "فئةٌ مقيسة", "من عددِ الوحدات"]

# المؤشّرُ الموسميُّ الغربيُّ المقيس — ٨ محطاتٍ لها سنتان فعليّتان كاملتان
# (وسيطُ ٢٠٢٥ و٢٠٢٦ لينا–سبت · وأكت–ديس من ٢٠٢٥ وحدها إذ لا فعليَّ لها في ٢٦)
SEAS_WEST = [1.049, 1.003, 1.013, 1.012, 0.939, 0.926,
             1.022, 1.053, 0.982, 0.952, 0.920, 0.882]
# أعمدةُ جدول المحطات في ورقة المُدخَلات
IC = dict(no=1, code=2, name=3, reg=4, city=5, open=6, off=7, day=8,
          basis=9, peer=10, clsr=11, built=12, uexp=13, urate=14, rate=15,
          src=16, n=17, cls=18, prof=19, m91=20, m95=21, mdsl=22)
MIXC = {"91": "m91", "95": "m95", "diesel": "mdsl"}
NIC = 22

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


def seas_ref(ir, m):
    """معاملُ موسمِ الشهر m للمحطة التي مُدخَلاتها في الصف ir

       يُختار المنحنى باسمه من عمودِ «الموسم»، ثم يُقسَم على متوسّطِ منحناه
       — فالقسمةُ تجعله رافعةَ شكلٍ لا رافعةَ مستوى: يُعيد توزيعَ اللترات
       على الشهور ولا يرفع مستوى المحطة.
    """
    lo, hi = R_PROF, R_PROF + N_PROF - 1
    key = f"'{IN}'!${L(IC['prof'])}${ir}"
    mt = f"MATCH({key},'{IN}'!$B${lo}:$B${hi},0)"
    cell = f"INDEX('{IN}'!$C${lo}:$N${hi},{mt},{m})"
    avg = f"INDEX('{IN}'!${L(C_PAVG)}${lo}:${L(C_PAVG)}${hi},{mt})"
    return f"{cell}/{avg}"


def ramp_factor(ir, m):
    """معاملُ التدرّج — وشهرُ الافتتاح بكسرِ أيامه لا بشهرٍ كامل

       شهرُ التشغيل الأولُ ليس شهراً كاملاً: محطةٌ تفتح في الخامسِ والعشرين
       تعمل سبعةَ أيامٍ من الشهر لا واحداً وثلاثين. فالوسيطُ المقيسُ ٠٫٢٩
       يُردُّ أولاً إلى أساسِ شهرٍ كاملٍ بقسمته على كسرِ عيّنةِ القياس،
       ثم يُضرَب في كسرِ أيامِ هذه المحطةِ بعينها.
    """
    off = f"'{IN}'!$G${ir}"
    day = f"'{IN}'!${L(IC['day'])}${ir}"
    opm = f"{m}-{off}"
    full = f"'{IN}'!$C${R_RAMP}/'{IN}'!${L(C_FRAC)}${R_RAMP}"
    frac = f"({days_ref(m)}-{day}+1)/{days_ref(m)}"
    later = f"INDEX({ramp_ref()},MIN(MAX({opm},1),7))"
    return f"IF({opm}=1,{full}*MAX(MIN({frac},1),0),{later})"


def fc_formula(ir, m, prod):
    """معادلةُ لترِ المنتجِ prod في الشهر m للمحطة التي مُدخَلاتها في الصف ir

       = معدّلُ النظير × التدرّج(شهرُ التشغيل) × أيامُ الشهر × حصّةُ المنتج

       وشهرُ التشغيل = m − «شهرُه»، حيث m شهرٌ ميلاديٌّ مرقَّمٌ من ١ و«شهرُه»
       موضعُ الافتتاح مرقَّمٌ من ٠ (٠ = يناير ٢٧ · −١ = ديسمبر ٢٦). فمحطةٌ
       تفتح في يناير ٢٧ شهرُها الأولُ هو m=١ ⇒ ١−٠ = ١. وما قبل الافتتاح صفر.
    """
    off = f"'{IN}'!$G${ir}"
    opm = f"{m}-{off}"
    rate = f"'{IN}'!${L(IC['rate'])}${ir}"
    mix = f"'{IN}'!${L(IC[MIXC[prod]])}${ir}"
    return (f"=IF({opm}<1,0,{rate}*{ramp_factor(ir, m)}"
            f"*{days_ref(m)}*{mix}*{seas_ref(ir, m)})")


# ══════════════════════════════════════════════════════════════════
def sheet_inputs(wb, D):
    ws = wb.create_sheet(IN)
    W.setup(ws, [5, 9, 24, 11, 13, 12, 6, 6, 15, 10, 10, 8, 8, 11, 11,
                 17, 5, 13, 13, 8, 8, 8], freeze="A23")
    W.title(ws, "المُدخَلات — ستّةُ مُدخَلاتٍ تُحرّك التوقّعَ كلَّه",
            "الأزرقُ يُغيَّر · والتوقّعُ في ورقة «FC 2027» معادلةٌ تقرأ من هنا، "
            "فما غيّرتَه هنا تحرّك هناك",
            NIC)
    W.rule(ws, 3, NIC)

    # ① منحنى التدرّج + كسرُ شهرِ الافتتاح
    c = ws.cell(4, 1, "① منحنى التدرّج — مقيسٌ من "
                      f"{D.get('n_ramp', 22)} افتتاحاً فعليّاً في ٢٠٢٦")
    c.font = Font(name=B.FONT, size=10, bold=True, color=B.BGRAY)
    ws.cell(R_RAMP, 2, "شهرُ التشغيل ←").font = Font(name=B.FONT, size=9, color=B.INK2)
    for k, v in enumerate(D["ramp"]):
        hc = ws.cell(R_RAMP - 1, 3 + k, f"شهر {k + 1}")
        hc.font = Font(name=B.FONT, size=8.5, bold=True, color="FFFFFF")
        hc.fill = PatternFill("solid", fgColor=B.BGRAY)
        hc.alignment = W.CTR
        W.inp(ws, R_RAMP, 3 + k, "0.00", round(v, 3))
    hc = ws.cell(R_RAMP - 1, C_FRAC, "كسرُ عيّنةِ\nالقياس")
    hc.font = Font(name=B.FONT, size=8, bold=True, color="FFFFFF")
    hc.fill = PatternFill("solid", fgColor=B.D_GOLD); hc.alignment = W.CTR
    W.inp(ws, R_RAMP, C_FRAC, "0.00", D.get("frac_sample", 0.5))
    ws.cell(R_RAMP, C_FRAC + 2,
            "شهرُ الافتتاح يُحسب بكسرِ أيامِه: ٠٫٢٩ ÷ كسرِ العيّنة = أساسُ "
            "شهرٍ كامل، ثم × (أيامُ الشهر − يومُ الافتتاح + ١) ÷ أيامِ الشهر"
            ).font = Font(name=B.FONT, size=8.5, italic=True, color=B.D_GOLD)

    # ② أيام الأشهر
    ws.cell(R_DAYS, 1, "② أيام ٢٠٢٧").font = Font(name=B.FONT, size=9.5, bold=True,
                                                   color=B.BGRAY)
    ws.cell(R_DAYS, 2, "فبراير ٢٨").font = Font(name=B.FONT, size=8.5, color=B.D_BAD)
    for m, d in enumerate(D["days"], 1):
        hc = ws.cell(R_DAYS - 1, 2 + m, MON[m - 1])
        hc.font = Font(name=B.FONT, size=8, bold=True, color="FFFFFF")
        hc.fill = PatternFill("solid", fgColor=B.BGRAY); hc.alignment = W.CTR
        W.inp(ws, R_DAYS, 2 + m, "0", d)

    # ③ منحنياتُ الموسم — يُختار منها واحدٌ لكلِّ محطةٍ في عمود «الموسم»
    ws.cell(R_PROF - 2, 1, "③ منحنياتُ الموسم — يُختار منها لكلِّ محطةٍ في "
                           "عمود «الموسم»، والقسمةُ على المتوسّطِ تجعله "
                           "رافعةَ شكلٍ لا مستوى")
    ws.cell(R_PROF - 2, 1).font = Font(name=B.FONT, size=10, bold=True,
                                       color=B.BGRAY)
    for m in range(1, 13):
        hc = ws.cell(R_PROF - 1, 2 + m, MON[m - 1])
        hc.font = Font(name=B.FONT, size=8, bold=True, color="FFFFFF")
        hc.fill = PatternFill("solid", fgColor=B.BGRAY); hc.alignment = W.CTR
    hc = ws.cell(R_PROF - 1, C_PAVG, "المتوسّط")
    hc.font = Font(name=B.FONT, size=8, bold=True, color="FFFFFF")
    hc.fill = PatternFill("solid", fgColor=B.BGRAY); hc.alignment = W.CTR
    sm = D["profiles"]["summer"]
    sn = D["profiles"]["summer_n"]
    PROFS = [("محايد", [1.0] * 12, "لا موسمَ — المستوى نفسُه كلَّ شهر"),
             ("غربيٌّ مقيس", SEAS_WEST,
              "٨ محطاتٍ غربيةٍ لها سنتان فعليّتان · قاعُ ماي–يون ٠٫٩٣ "
              "وقمّةُ أغس ١٫٠٥"),
             ("صيفيٌّ مقيس", sm,
              f"أعلى {sn} محطاتٍ صيفيّةً · مقيسٌ ينا–سبت، وأكت–ديس على "
              f"مستوى ينا–أبر افتراضاً"),
             ("مخصَّص", [1.0] * 12, "فارغٌ لتكتب فيه منحنى موقعٍ بعينه")]
    for j, (nm, vals, note) in enumerate(PROFS):
        rr = R_PROF + j
        cc = ws.cell(rr, 2, nm)
        cc.font = Font(name=B.FONT, size=9, bold=True, color=B.BGRAY)
        cc.border = W.BOX; cc.alignment = W.CTR; cc.fill = NEU
        for m, v in enumerate(vals, 1):
            W.inp(ws, rr, 2 + m, "0.000", round(v, 3))
        W.calc(ws, rr, C_PAVG,
               f"=AVERAGE({L(3)}{rr}:{L(14)}{rr})", "0.000")
        cc = ws.cell(rr, C_PAVG + 1, note)
        cc.font = Font(name=B.FONT, size=8.5, italic=True, color=B.INK2)
        cc.alignment = W.RGT

    # ④ رافعةُ المستأجرين — مقيسةٌ على ٦٧ محطةً ناضجةً لها سجلُّ وحدات
    a = D.get("anchor") or {}
    ws.cell(R_ANC - 2, 1,
            f"④ رافعةُ المستأجرين — عددُ الوحداتِ المؤجَّرة يفسّر "
            f"{a.get('r2', 0) * 100:.0f}٪ من تباينِ اللتر (ن={a.get('n', 0)})، "
            f"والعلامةُ الكبرى {a.get('r2_brands', 0) * 100:.0f}٪ فقط")
    ws.cell(R_ANC - 2, 1).font = Font(name=B.FONT, size=10, bold=True,
                                      color=B.BGRAY)
    for j, t in enumerate(("الميل\nلتر/وحدة", "الثابت\nلتر/يوم",
                           "إشغالٌ\nمفترض", "السقف\nل/يوم",
                           "وسيطُ وحداتِ\nالشبكة")):
        hc = ws.cell(R_ANC - 1, 3 + j, t)
        hc.font = Font(name=B.FONT, size=8, bold=True, color="FFFFFF")
        hc.fill = PatternFill("solid", fgColor=B.D_GOLD); hc.alignment = W.CTR
    W.inp(ws, R_ANC, C_SLOPE, NUM, a.get("slope", 0))
    W.inp(ws, R_ANC, C_INTER, NUM, a.get("intercept", 0))
    W.inp(ws, R_ANC, C_OCC, "0.0%", a.get("occ", 0))
    W.inp(ws, R_ANC, C_CAP, NUM, a.get("cap_rate", 0))
    W.calc(ws, R_ANC, 7, a.get("net_units", 0), "0.0")
    ws.cell(R_ANC, 8,
            f"السقفُ = أعلى محطةٍ قائمة ({a.get('cap_name', '')}) — فلا يُوازَن "
            f"موقعٌ فوق ما لم نشغّله قطّ · والإشغالُ مجموعُ المؤجَّرِ ÷ المبنيِّ "
            f"في الناضجة").font = Font(name=B.FONT, size=8.5, italic=True,
                                        color=B.INK2)

    # ⑤ جدول المحطات
    r = R_ANC + 2
    ws.cell(r, 1, "⑤ لكل محطة: أساسُ معدّلها — ثلاثةُ أسُسٍ مقيسةٍ تُقلَّب "
                  "من القائمةِ المنسدلة، والمعدّلُ يُحسب من المختار")
    ws.cell(r, 1).font = Font(name=B.FONT, size=10, bold=True, color=B.BGRAY)
    W.header(ws, R_ST - 1, ["#", "الكود", "المحطة", "المنطقة", "المدينة",
                            "الافتتاح", "شهرُه*", "يومُه",
                            "أساسُ المعدّل", "① نظيرُ\nالجوار",
                            "② فئةٌ\nمقيسة", "وحداتٌ\nمبنيّة",
                            "وحداتٌ\nمتوقَّعة", "③ من عددِ\nالوحدات",
                            "المعدّلُ\nالمختار", "مصدرُ النظير", "ن",
                            "الفئة", "الموسم", "٩١", "٩٥", "ديزل"], h=34)
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
        W.inp(ws, rr, IC["day"], "0", s["day"])
        cc = W.inp(ws, rr, IC["basis"], "@", s.get("basis") or BASES[0])
        if (s.get("basis") or BASES[0]) != BASES[0]:
            cc.font = Font(name=B.FONT, size=9, bold=True, color=B.D_GOLD)
        W.inp(ws, rr, IC["peer"], NUM, s.get("peer_rate", s["rate"]))
        cr = s.get("cls_rate", 0)
        cc = W.inp(ws, rr, IC["clsr"], NUM, cr or None)
        if not cr:
            cc.value = None
        W.inp(ws, rr, IC["built"], "0", s.get("units_built", 0))
        W.calc(ws, rr, IC["uexp"],
               f"=ROUND({L(IC['built'])}{rr}*'{IN}'!${L(C_OCC)}${R_ANC},0)", "0")
        W.calc(ws, rr, IC["urate"],
               f"=IF({L(IC['built'])}{rr}=0,\"\",MIN("
               f"'{IN}'!${L(C_INTER)}${R_ANC}+'{IN}'!${L(C_SLOPE)}${R_ANC}"
               f"*{L(IC['uexp'])}{rr},'{IN}'!${L(C_CAP)}${R_ANC}))", NUM)
        W.calc(ws, rr, IC["rate"],
               f'=IF({L(IC["basis"])}{rr}="{BASES[1]}",{L(IC["clsr"])}{rr},'
               f'IF(AND({L(IC["basis"])}{rr}="{BASES[2]}",'
               f'{L(IC["built"])}{rr}>0),{L(IC["urate"])}{rr},'
               f'{L(IC["peer"])}{rr}))', NUM)
        cc = ws.cell(rr, IC["src"], s["src"]); cc.font = Font(name=B.FONT, size=8.5,
                                                              color=B.INK2)
        cc.border = W.BOX; cc.alignment = W.RGT; cc.fill = NEU
        W.calc(ws, rr, IC["n"], s["n"], "0")
        cc = W.inp(ws, rr, IC["cls"], "@", s.get("cls") or "—")
        if s.get("cls"):
            cc.font = Font(name=B.FONT, size=9, bold=True, color=B.D_GOLD)
        cc = W.inp(ws, rr, IC["prof"], "@", s.get("prof") or "محايد")
        if (s.get("prof") or "محايد") != "محايد":
            cc.font = Font(name=B.FONT, size=9, bold=True, color=B.D_GOLD)
        for p in ("91", "95", "diesel"):
            W.inp(ws, rr, IC[MIXC[p]], "0.0%", s["mix"][p])
    last = R_ST + len(D["stations"]) - 1
    for col, opts in ((IC["prof"], [p[0] for p in PROFS]),
                      (IC["basis"], BASES),
                      (IC["cls"], ["—"] + list(D.get("classes", {})))):
        dv = DataValidation(type="list",
                            formula1='"' + ",".join(opts) + '"',
                            allow_blank=False)
        ws.add_data_validation(dv)
        dv.add(f"{L(col)}{R_ST}:{L(col)}{last}")
    # المعدّلُ المختارُ فوق سقفِ الشبكة — تحذيرٌ بصريّ
    ws.conditional_formatting.add(
        f"{L(IC['rate'])}{R_ST}:{L(IC['rate'])}{last}",
        CellIsRule(operator="greaterThan",
                   formula=[f"'{IN}'!${L(C_CAP)}${R_ANC}"], fill=BAD,
                   font=Font(name=B.FONT, size=10, bold=True, color=B.D_BAD)))
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
    NOTES = [
        ("⚠", B.D_GOLD,
         "عمودا «الفئة» و«الموسم» هما ما يُخرج المحطةَ من وسيطِ جيرانها. "
         "والمذهَّبُ منهما مُدخَلٌ إداريٌّ لا قياسٌ مستخرَج — مصدرُه مكتوبٌ "
         "في ورقة «ما ينقص»."),
        ("•", B.INK2,
         "«ممرُّ شاحنات» فئةٌ مُنتزعةٌ من الناضجةِ بحصّةِ الديزل ٤٠–٥٥٪، لا "
         "بالمقاسِ المسجَّل: كلُّ الـ٣٨ مسجَّلةٌ Small ومنها ما يفوق مئةَ "
         "ألفِ متر، ووسيطُ الـMedium القائمةِ يتشتَّت من ١٥٬٦٦٢ إلى ٩٦٬٨٨٠ "
         "— فالمقاسُ لا يفرز، وحصّةُ الديزل تفرز."),
        ("•", B.INK2,
         "«غربيٌّ مقيس» هو الموسمُ الافتراضيُّ لمحطاتِ الغربية وحدها، لأنّه "
         "قيس على ثماني محطاتٍ غربيةٍ لها سنتان فعليّتان: قاعُ مايو–يونيو "
         "٠٫٩٣ وقمّةُ أغسطس ١٫٠٥، وقاعُه شهرُ عيد الأضحى في السنتين "
         "(٦ يونيو ٢٠٢٥ · ٢٧ مايو ٢٠٢٦)."),
        ("•", B.INK2,
         "وما عدا الغربية «محايد» — لأنّ موسمَها غيرُ مقيسٍ لا لأنّه معدوم. "
         "ولا يُملأ بمنحنى الغربية، فالحجُّ والعمرةُ يحرّكان مكةَ ولا يحرّكان "
         "جازانَ ولا حائلَ."),
        ("•", B.INK2,
         "وفي ٢٠٢٧ تتقدّم المواقيتُ نحو أحدَ عشرَ يوماً: رمضان ٨ فبراير–٩ "
         "مارس · عيد الفطر ٩–١٠ مارس · عيد الأضحى ١٦–١٧ مايو. فالقاعُ "
         "المقيسُ في مايو يوافق ٢٠٢٧، والقاعُ المقيسُ في يونيو لا يوافقها."),
        ("•", B.INK2,
         "والقسمةُ على متوسّطِ المنحنى تجعله رافعةَ شكلٍ لا مستوى: محطةٌ "
         "تعمل السنةَ كاملةً يتغيّر توزيعُها الشهريُّ ويثبت إجماليُّها، "
         "ومحطةٌ تفتح في أثناء السنة يتحرّك إجماليُّها بقدرِ موسمِ الأشهر "
         "التي عملتها وحدها."),
    ]
    for j, (bullet, col, t) in enumerate(NOTES):
        cc = ws.cell(last + 5 + j, 1, f"{bullet}  {t}")
        cc.font = Font(name=B.FONT, size=8.5, italic=True, color=col)
        ws.merge_cells(start_row=last + 5 + j, start_column=1,
                       end_row=last + 5 + j, end_column=NIC)
        cc.alignment = W.RGT
    last += 5 + len(NOTES)

    # جداولُ مرجعيّة
    r = last + 5
    r = W.band(ws, r, NIC, "وسيطُ النظراء كما قيس — ينا–سبت ٢٠٢٦ الفعليّ")
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

    r = W.band(ws, r, NIC, "مزيجُ المنتجات الفعليُّ بالمنطقة — ينا–سبت ٢٠٢٦")
    W.header(ws, r, ["المنطقة", "٩١", "٩٥", "ديزل"], h=20)
    r += 1
    for rg, m in D["mix_reg"].items():
        W.calc(ws, r, 1, rg, "@")
        for j, p in enumerate(("91", "95", "diesel")):
            W.calc(ws, r, 2 + j, m[p], "0.0%")
        r += 1
    r += 1

    # الفئاتُ المقيسة — بأعضائها، فالوسيطُ بلا أعضائه لا يُراجَع
    r = W.band(ws, r, NIC, "الفئاتُ المقيسة — مُنتزعةً من الناضجةِ بحصّةِ "
                           "الديزل، وبأعضائها كي تُراجَع")
    W.header(ws, r, ["الفئة", "ل/يوم وسيطاً", "ر٣", "الأعلى", "ن",
                     "٩١", "٩٥", "ديزل", "الأعضاء"], h=22)
    r += 1
    for nm, c in D.get("classes", {}).items():
        W.calc(ws, r, 1, nm, "@")
        W.calc(ws, r, 2, c["rate"], NUM)
        W.calc(ws, r, 3, c["q3"], NUM)
        W.calc(ws, r, 4, c["top"], NUM)
        W.calc(ws, r, 5, c["n"], "0")
        for j, p in enumerate(("91", "95", "diesel")):
            W.calc(ws, r, 6 + j, c["mix"][p], "0%")
        cc = ws.cell(r, 9, " · ".join(f"{n} {v:,}" for n, v, _ in c["members"]))
        cc.font = Font(name=B.FONT, size=8, color=B.INK2)
        cc.border = W.BOX; cc.alignment = W.RGT
        r += 1
    r += 1
    # رافعةُ المستأجرين — الدليلُ والنقيضُ معاً
    a = D.get("anchor")
    if a:
        r += 1
        r = W.band(ws, r, NIC,
                   f"رافعةُ المستأجرين كما قيست — {a['n']} محطةً ناضجةً لها "
                   f"سجلُّ وحدات · والنقيضُ مكتوبٌ معها لأنّ السببيّةَ "
                   f"غيرُ مُثبَتة")
        W.header(ws, r, ["وحداتٌ مؤجَّرة", "ن", "وسيطُ اللتر", "معامل"], h=20)
        r += 1
        for lbl, n, med, fac in a["buckets"]:
            W.calc(ws, r, 1, lbl, "@"); W.calc(ws, r, 2, n, "0")
            W.calc(ws, r, 3, med, NUM)
            W.calc(ws, r, 4, fac, '0.00"×"'); r += 1
        for ttl, items, col in (
                ("يرفع الرافعة", [f"ارتباطُ الوحداتِ المؤجَّرة r={a['r_leased']:+.3f}"
                                  f" (R²={a['r2']:.2f}) — أقوى ما قيس في هذا "
                                  f"الملف",
                                  f"والعلاقةُ متدرّجة: من {a['buckets'][0][2]:,} "
                                  f"ل/يوم بلا مستأجرٍ إلى "
                                  f"{a['buckets'][-1][2]:,} عند أربعةَ عشرَ "
                                  f"مستأجراً وأكثر"], B.D_GOOD),
                ("ويهدم سببيّتها",
                 [f"{len(a['zero_high'])} محطاتٍ بلا مستأجرٍ واحدٍ تبيع فوق "
                  f"٤٠ ألف ل/يوم: "
                  + " · ".join(f"{n} {v:,}" for n, v in a["zero_high"])]
                 + [f"و{len(a['many_low'])} بعشرِ وحداتٍ مؤجَّرةٍ أو أكثرَ تبيع "
                    f"دون ٢٠ ألف: "
                    + " · ".join(f"{n} {v:,} ({u} وحدة · {b} علامة)"
                                 for n, v, u, b in a["many_low"])]
                 + [f"والعلامةُ الكبرى وحدَها تفسّر {a['r2_brands'] * 100:.0f}٪ "
                    f"فقط — فالعددُ يفسّر لا الاسم: "
                    + " · ".join(f"{n} ({b} علامة) {v:,} ل/يوم"
                                 for n, b, u, v in a["richest"][:3])]
                 + [f"و{a['max_units']} وحدةً أعلى ما شغّلناه، فما فوقها "
                    f"استقراءٌ خارجَ البيانات — والسقفُ "
                    f"{a['cap_rate']:,} ل/يوم يمنع تجاوزَ "
                    f"{a['cap_name']}"], B.D_BAD)):
            cc = ws.cell(r, 1, ttl)
            cc.font = Font(name=B.FONT, size=9, bold=True, color=col)
            cc.alignment = W.RGT
            r += 1
            for it in items:
                ws.merge_cells(start_row=r, start_column=2, end_row=r,
                               end_column=NIC)
                cc = ws.cell(r, 2, "•  " + it)
                cc.font = Font(name=B.FONT, size=8.5, color=B.INK2)
                cc.alignment = W.RGT
                r += 1
        r += 1

    sm = D["profiles"]
    r = W.band(ws, r, NIC, f"أعضاءُ منحنى «صيفيٌّ مقيس» — أعلى "
                           f"{sm['summer_n']} محطاتٍ في نسبةِ (يون–أغس ÷ "
                           f"ينا–ماي+سبت) · ووسيطُ الشبكة "
                           f"{sm['summer_median']:.2f}×")
    W.header(ws, r, ["المحطة", "المدينة", "نسبةُ الصيف"], h=20)
    r += 1
    for nm, ct, ratio in sm["summer_members"]:
        W.calc(ws, r, 1, nm, "@"); W.calc(ws, r, 2, ct, "@")
        W.calc(ws, r, 3, ratio, "0.00\"×\""); r += 1
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
    W.setup(ws, [4, 26, 11, 12, 8, 13, 13, 11, 11, 12, 12, 11, 11, 24],
            freeze="A5")
    W.title(ws, "الفجوة — أين يتركّز الفارقُ بين الموازنة والتوقّع",
            "مرتَّبةً بالموازنة · وعمودُ «ل/يوم تُوازِنه» مقابلَ أعلى محطةٍ "
            f"قائمةٍ ({D.get('top_name', '')} {D.get('top_mature', 0):,} ل/يوم) "
            "هو اختبارُ كلِّ سطر",
            14)
    W.rule(ws, 3, 14)
    W.header(ws, 4, ["#", "المحطة", "الكود", "الافتتاح", "أشهرُ\nالتشغيل",
                     "الفئة", "الموسم",
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
        cc = W.calc(ws, r, 6, s.get("cls") or "—", "@")
        if s.get("cls"):
            cc.font = Font(name=B.FONT, size=9, bold=True, color=B.D_GOLD)
        cc = W.calc(ws, r, 7, s.get("prof") or "محايد", "@")
        if (s.get("prof") or "محايد") != "محايد":
            cc.font = Font(name=B.FONT, size=9, bold=True, color=B.D_GOLD)
        W.calc(ws, r, 8, bu / 1e6, "0.00")
        W.calc(ws, r, 9, f"='FC 2027'!U{f}/1000000", "0.00")
        W.calc(ws, r, 10, f"={L(9)}{r}-{L(8)}{r}", "+0.00;-0.00")
        W.calc(ws, r, 11, f"=IF({L(8)}{r}=0,\"\",{L(9)}{r}/{L(8)}{r})", "0%")
        W.calc(ws, r, 12, f"='{IN}'!${L(IC['rate'])}${R_ST + i - 1}", NUM)
        W.calc(ws, r, 13, rt_bu, NUM)
        top = D.get("top_mature", 0)
        if s["bu_flat"]:
            note = (f"موازنتُها رقمٌ واحدٌ مكرَّرٌ في {mo} شهراً — "
                    "بلا تدرّجٍ ولا موسم")
        elif rt_bu > top:
            note = (f"تُوازَن فوق أعلى محطةٍ قائمة ({rt_bu / top:.2f}×) — "
                    "تحتاج عدّادَ حركةٍ لا قالباً")
        elif s["n"] < 5:
            note = f"نظراءُ {s['n']} فقط — الأساسُ هشّ"
        else:
            note = "الحركةُ المارّةُ بالموقع — وسيطُ المدينة لا يَحسمها"
        cc = ws.cell(r, 14, note)
        cc.font = Font(name=B.FONT, size=8.5,
                       color=B.D_BAD if (rt_bu > top or s["bu_flat"]) else B.INK2)
        cc.border = W.BOX; cc.alignment = W.RGT
        r += 1
    tot = r
    cc = ws.cell(r, 2, "المجموع — ٣٨ محطة")
    cc.font = Font(name=B.FONT, size=10, bold=True, color="FFFFFF")
    cc.fill = PatternFill("solid", fgColor=B.BGRAY); cc.alignment = W.RGT
    cc.border = W.BOX
    for cnum, f, fmt in ((8, f"=SUM(H5:H{r - 1})", "0.0"),
                         (9, f"=SUM(I5:I{r - 1})", "0.0"),
                         (10, f"=I{r}-H{r}", "+0.0;-0.0"),
                         (11, f"=I{r}/H{r}", "0%")):
        cc = W.calc(ws, r, cnum, f, fmt)
        cc.fill = BAND; cc.font = Font(name=B.FONT, size=10, bold=True, color=B.INK)
    for op, fl, col in (("lessThan", BAD, B.D_BAD),
                        ("greaterThan", GOOD, B.D_GOOD)):
        ws.conditional_formatting.add(f"J5:J{tot}", CellIsRule(
            operator=op, formula=["0"], fill=fl,
            font=Font(name=B.FONT, size=9.5, bold=True, color=col)))
    ws.conditional_formatting.add(f"K5:K{tot - 1}", ColorScaleRule(
        start_type="num", start_value=0, start_color=B.T_BAD,
        mid_type="num", mid_value=1, mid_color="FFFFFF",
        end_type="num", end_value=2, end_color=B.T_GOOD))
    ws.conditional_formatting.add(f"M5:M{tot - 1}", CellIsRule(
        operator="greaterThan", formula=[str(D.get("top_mature", 90000))],
        fill=BAD, font=Font(name=B.FONT, size=9.5, bold=True, color=B.D_BAD)))
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
        ("خطّةُ مستأجري القنفذة ومدينة الملك عبدالله — عدداً وعلامةً وموعداً",
         "٦٥ وحدةً مبنيّةً في القنفذة و٤١ في مدينة الملك، ومعدّلُهما محسوبٌ "
         "بإشغالِ ٧٧٫٥٪ (إشغالُ الناضجة مجموعاً). والمحلاتُ وحدَها في الشبكة "
         "٢٥٫١٪ إشغالاً، فالفرقُ بين الفرضَين يحرّك الموقعَ أكثرَ من أيِّ "
         "مُدخَلٍ آخر",
         "إدارةُ العقار",
         "يصحّح «وحداتٌ مبنيّة» و«إشغالٌ مفترض» فيتحرّك معدّلُ المحطتين"),
        ("أيامُ افتتاحِ الـ٢٢ التي قيس بها التدرّج",
         "تواريخُها غيرُ مسجَّلةٍ في الملف، فكسرُ شهرِها الأولِ مفترضٌ ٠٫٥٠ "
         "(توزُّعٌ منتظمٌ داخلَ الشهر) لا مقيساً",
         "إدارةُ المشاريع",
         "يصحّح خليّةَ «كسرُ عيّنةِ القياس» فيتحرّك شهرُ الافتتاح لكلِّ الـ٣٨"),
        ("الحركةُ المارّةُ بكل موقعٍ جديد (زيارة/يوم)",
         "الوحدةُ تبيع ما يمرّ بها — ووسيطُ المدينة لا يعرف طريقَ المحطة",
         "هندسةُ الشبكة",
         "يستبدل العمودَ «معدّل النظير» كلَّه بقياسٍ موقعيّ"),
        ("أشهرٌ فعليّةٌ كاملةٌ لـ٢٠٢٥ لغير المنطقة الغربية",
         "الثماني الوحيدةُ التي لها سنتان كلُّها غربية — فـ١٦ محطةً جنوبيةً "
         "ووسطى وشماليةً موسمُها «محايد» لأنّه غيرُ مقيس، لا لأنّه معدوم",
         "الإدارة المالية",
         "يُنشئ منحنىً لكلِّ منطقةٍ بدل «محايد»"),
        ("أكتوبر–ديسمبر لمنحنى «صيفيٌّ مقيس»",
         "لا فعليَّ لها في ٢٠٢٦، فوُضعت على مستوى ينا–أبر المقيسِ افتراضاً",
         "الإدارة المالية",
         "يصحّح ثلاثةَ أشهرٍ في منحنى الصيف — وأثرُها على مدينة الملك عبدالله"),
        ("عدّادُ الشاحناتِ على ممرِّ القنفذة ومدينة الملك عبدالله",
         "فئةُ «ممرِّ شاحنات» وسيطُها ٣٢٬٨٠٩ وتتشتَّت إلى ٥٤٬٥٢٦ على ستِّ "
         "محطاتٍ فقط — والعدّادُ يحسم أيَّ موضعٍ من المدى",
         "هندسةُ الشبكة",
         "يستبدل وسيطَ الفئة بمعدّلٍ موقعيٍّ للمحطتين"),
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

    # ما تعرفه الإدارةُ عن الموقع ولا يَظهر في الملف — مُدخَلٌ إداريٌّ مُعلَن
    ov = [s for s in D["stations"] if s.get("why")]
    if ov:
        r = W.band(ws, r, 5, "مُدخَلاتٌ إداريّة — معرفةُ موقعٍ لا يحملها الملف، "
                             "وليست قياساً مستخرَجاً")
        W.header(ws, r, ["#", "المحطة", "ما قيل عن الموقع", "ما غُيِّر في النموذج",
                         "المصدر"], h=24)
        r += 1
        for i, s in enumerate(ov, 1):
            W.calc(ws, r, 1, i, "0")
            why = s["why"].split(" — ")
            chg = []
            if s.get("cls"):
                chg.append(f"الفئة ← «{s['cls']}» ({s['rate']:,} ل/يوم "
                           f"بدل وسيطِ الجوار)")
            if (s.get("prof") or "محايد") != "محايد":
                chg.append(f"الموسم ← «{s['prof']}»")
            for cnum, v, col, bold in ((2, s["name"], B.INK, True),
                                       (3, why[0], B.INK2, False),
                                       (4, " · ".join(chg), B.D_GOLD, False),
                                       (5, why[-1] if len(why) > 1 else "",
                                        B.INK2, False)):
                cc = ws.cell(r, cnum, v)
                cc.font = Font(name=B.FONT, size=9, bold=bold, color=col)
                cc.border = W.BOX; cc.alignment = W.WRAP
            ws.row_dimensions[r].height = 40
            r += 1
        r += 1

    W.bullets(ws, r, 5, [(
        "كيف تُقرأ هذه الورقة",
        ["التوقّعُ هنا أفضلُ قراءةٍ تحتملها بياناتنا — لا قياسُ موقع",
         "وخمسةُ مُدخَلاته مقيسة: معدّلُ النظير · منحنى التدرّج · كسرُ شهرِ "
         "الافتتاح · منحنى الموسم · مزيجُ المنتجات",
         "والمُدخَلُ الإداريُّ مفصولٌ في الجدول أعلاه بمصدره — فلا يُحسب قياساً",
         "وما ليس هذا ولا ذاك مكتوبٌ بمالكه، ولم يُملأ برقمٍ يبدو قياساً"],
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
