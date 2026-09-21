# -*- coding: utf-8 -*-
"""ملف حافز المحطات — إعادة بناء منظَّمة لملفَّي MK019 وMK072

   ما صُحِّح عن الملف الأصلي:
   ① النسبة كانت تقارن فعلياً لتسعة عشر يوماً بميزانية ثلاثين — فتخرج ٥٩٪
      وتُسقط الحافز. الصحيح ٩٢٫٥٪ و٩٣٫١٪، وكلاهما داخل شريحة الصرف.
   ② العملة في بطاقة الأداء كانت JOD — قالبٌ أردني لم يُعرَّب ولم يُحوَّل.
   ③ الأرقام كانت مطبوعة لا محسوبة؛ هنا كل خلية نتيجةٍ صيغةٌ حيّة.
   ④ ملف لكل محطة؛ هنا ملف واحد يقبل أي عدد من المحطات.
   ⑤ ورقة يومية ٢٬١٥٣ صفاً × ١٢٩ عموداً وفيها ثلاثون رقماً — هنا جدول بحجمه.

   يُشغَّل من جذر المستودع:  PYTHONPATH=data python3 data/build_incentive.py
"""
import json, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule, DataBarRule
from openpyxl.worksheet.datavalidation import DataValidation

import brand as B
import worker_model as W

SRC = "data/incentive-source.json"
OUT = "docs/حافز-المحطات.xlsx"
MONTH_AR = {"01": "يناير", "02": "فبراير", "03": "مارس", "04": "أبريل", "05": "مايو",
            "06": "يونيو", "07": "يوليو", "08": "أغسطس", "09": "سبتمبر", "10": "أكتوبر",
            "11": "نوفمبر", "12": "ديسمبر"}
DAYS_IN_MONTH = 30
DAYS_ELAPSED = 19          # أيام سبتمبر التي وصلت بياناتها — يُعدَّل فتتحدّث الورقة كلها

MONEY = '#,##0;[Red]-#,##0'
MONEY2 = '#,##0.00;[Red]-#,##0.00'
PCT = '0.0%'


def arn(n):
    """أرقام عربية-هندية — للنصوص وحدها، أمّا خلايا الحساب فتبقى لاتينية"""
    return str(n).translate(str.maketrans("0123456789", "٠١٢٣٤٥٦٧٨٩"))


def sheet(wb, name, widths, title, sub, nc, idx=None):
    ws = wb.create_sheet(name) if idx is None else wb.create_sheet(name, idx)
    W.setup(ws, widths, freeze="A4")
    W.title(ws, title, sub, nc)
    return ws


def cell(ws, r, c, v, *, fmt=None, font=None, fill=None, align=None, box=True):
    x = ws.cell(r, c, v)
    if fmt: x.number_format = fmt
    x.font = font or W.BLACK
    if fill: x.fill = fill
    x.alignment = align or W.CTR
    if box: x.border = W.BOX
    return x


# ═══════════════════════════════════════════════════════ ① لوحة الحافز
def dashboard(wb, S, idx=0):
    NC = 11
    ws = sheet(wb, "لوحة الحافز", [4, 20, 9, 8, 13, 13, 8, 13, 12, 12, 30],
               "حافز المحطات — سبتمبر ٢٠٢٦",
               f"النسبة تُقاس على الأيام التي وصلت بياناتها ({arn(DAYS_ELAPSED)} يوماً من "
               f"{arn(DAYS_IN_MONTH)}) — لا على الشهر كاملاً", NC, idx)
    r = W.band(ws, 3, NC, "① النتيجة — ولكل محطة سطر واحد")
    W.header(ws, r, ["#", "المحطة", "الكود", "الطاقم", "ميزانية الشهر",
                     f"ميزانية {arn(DAYS_ELAPSED)} يوماً", "الفعلي", "التحقيق",
                     "الشريحة", "وعاء الحافز", "القراءة"])
    r += 1
    first = r
    for i, s in enumerate(S, 1):
        pot = tier_amount(s)
        ach = s["act"] / (s["bu"] / DAYS_IN_MONTH * DAYS_ELAPSED)
        cell(ws, r, 1, i)
        cell(ws, r, 2, s["name"], font=W.BOLD, align=W.RGT)
        cell(ws, r, 3, s["code"])
        cell(ws, r, 4, s["staff"])
        cell(ws, r, 5, s["bu"], fmt=MONEY)
        cell(ws, r, 6, f"=E{r}/الاحتساب!$C$5*الاحتساب!$C$6", fmt=MONEY,
             fill=PatternFill("solid", fgColor=B.T_NEUTRAL))
        cell(ws, r, 7, s["act"], fmt=MONEY, font=W.BOLD)
        c = cell(ws, r, 8, f"=IFERROR(G{r}/F{r},0)", fmt=PCT, font=W.BOLD,
                 fill=PatternFill("solid", fgColor=B.T_NEUTRAL))
        cell(ws, r, 9, tier_name(ach), font=W.BOLD)
        cell(ws, r, 10, pot, fmt=MONEY, font=W.BOLD)
        cell(ws, r, 11, "يُصرف" if ach >= 0.85 else "دون العتبة — لا يُصرف",
             align=W.WRAP, font=W.BLACK)
        ws.row_dimensions[r].height = 20
        r += 1
    last = r - 1
    ws.conditional_formatting.add(f"H{first}:H{last}", CellIsRule(
        operator="greaterThanOrEqual", formula=["0.85"],
        fill=PatternFill("solid", fgColor=B.T_GOOD), font=Font(name=B.FONT, bold=True, color=B.GOOD)))
    ws.conditional_formatting.add(f"H{first}:H{last}", CellIsRule(
        operator="lessThan", formula=["0.85"],
        fill=PatternFill("solid", fgColor=B.T_BAD), font=Font(name=B.FONT, bold=True, color=B.BAD)))

    r = last + 2
    r = W.band(ws, r, NC, "② ما كان يُحتسب — ولماذا كان يُسقط الحافز")
    W.header(ws, r, ["#", "المحطة", "الكود", "", "النسبة في الملف الأصلي", "الصحيحة",
                     "الفارق", "الأثر", "", "", "الشرح"])
    r += 1
    for i, s in enumerate(S, 1):
        old = s["act"] / s["bu"]
        new = s["act"] / (s["bu"] / DAYS_IN_MONTH * DAYS_ELAPSED)
        cell(ws, r, 1, i)
        cell(ws, r, 2, s["name"], font=W.BOLD, align=W.RGT)
        cell(ws, r, 3, s["code"])
        cell(ws, r, 4, "")
        cell(ws, r, 5, old, fmt=PCT, fill=PatternFill("solid", fgColor=B.T_BAD),
             font=Font(name=B.FONT, size=10, bold=True, color=B.BAD))
        cell(ws, r, 6, new, fmt=PCT, fill=PatternFill("solid", fgColor=B.T_GOOD),
             font=Font(name=B.FONT, size=10, bold=True, color=B.GOOD))
        cell(ws, r, 7, new - old, fmt=PCT, font=W.BOLD)
        cell(ws, r, 8, tier_amount(s), fmt=MONEY, font=W.BOLD)
        cell(ws, r, 9, ""); cell(ws, r, 10, "")
        cell(ws, r, 11, f"الفعلي {arn(DAYS_ELAPSED)} يوماً قُسم على ميزانية {arn(DAYS_IN_MONTH)} يوماً",
             align=W.WRAP)
        ws.row_dimensions[r].height = 20
        r += 1
    ws.cell(r + 1, 2, "الأثر: محطتان تُقرآن «دون العتبة» وهما فوقها — "
                      "والحافز المستحق يسقط بلا سبب.").font = W.SMALL
    ws.cell(r + 1, 2).alignment = W.RGT
    return ws


def tier_name(ach):
    if ach >= 1.0: return "١٠٠٪ فأكثر"
    if ach >= 0.85: return "٨٥٪ إلى ٩٩٪"
    return "دون ٨٥٪"


def tier_amount(s):
    ach = s["act"] / (s["bu"] / DAYS_IN_MONTH * DAYS_ELAPSED)
    t = s["tiers"]
    if ach >= 1.0: return t[2]["pot"]
    if ach >= 0.85: return t[0]["pot"]
    return 0


# ═══════════════════════════════════════════════════════ ② الاحتساب
def calc(wb, S, idx=1):
    NC = 9
    ws = sheet(wb, "الاحتساب", [4, 26, 15, 15, 14, 14, 12, 12, 34],
               "الاحتساب — كل رقم صيغة لا قيمة مطبوعة",
               "غيِّر الأصفر فيتحدّث ما تحته · وهذه الورقة هي مصدر لوحة الحافز", NC, idx)
    r = W.band(ws, 3, NC, "① المُدخَلات — تُعدَّل هنا وحدها")
    IN = PatternFill("solid", fgColor=B.T_GOLD)
    cell(ws, r + 1, 2, "أيام الشهر", font=W.BOLD, align=W.RGT)
    cell(ws, r + 1, 3, DAYS_IN_MONTH, fill=IN, font=W.BOLD)
    cell(ws, r + 2, 2, "الأيام التي وصلت بياناتها", font=W.BOLD, align=W.RGT)
    cell(ws, r + 2, 3, DAYS_ELAPSED, fill=IN, font=W.BOLD)
    cell(ws, r + 3, 2, "عتبة الصرف", font=W.BOLD, align=W.RGT)
    cell(ws, r + 3, 3, 0.85, fmt=PCT, fill=IN, font=W.BOLD)
    cell(ws, r + 4, 2, "عتبة الشريحة العليا", font=W.BOLD, align=W.RGT)
    cell(ws, r + 4, 3, 1.00, fmt=PCT, fill=IN, font=W.BOLD)
    ws.cell(r + 1, 5, "يُقاس التحقيق على الأيام المتاحة لا على الشهر كاملاً — "
                      "وإلا قُورن جزءٌ بكلٍّ.").font = W.SMALL
    ws.cell(r + 1, 5).alignment = W.RGT

    r = W.band(ws, r + 6, NC, "② الاحتساب لكل محطة")
    W.header(ws, r, ["#", "المحطة", "ميزانية الشهر", f"ميزانية الأيام المتاحة",
                     "الفعلي", "التحقيق", "الوعاء", "حصة المدير", "حصة العامل الواحد"])
    r += 1
    for i, s in enumerate(S, 1):
        t = s["tiers"]
        cell(ws, r, 1, i)
        cell(ws, r, 2, f"{s['name']} · {s['code']}", font=W.BOLD, align=W.RGT)
        cell(ws, r, 3, s["bu"], fmt=MONEY)
        cell(ws, r, 4, f"=C{r}/$C$5*$C$6", fmt=MONEY,
             fill=PatternFill("solid", fgColor=B.T_NEUTRAL))
        cell(ws, r, 5, s["act"], fmt=MONEY, font=W.BOLD)
        cell(ws, r, 6, f"=IFERROR(E{r}/D{r},0)", fmt=PCT, font=W.BOLD,
             fill=PatternFill("solid", fgColor=B.T_NEUTRAL))
        cell(ws, r, 7, f'=IF(F{r}>=$C$8,{t[2]["pot"]},IF(F{r}>=$C$7,{t[0]["pot"]},0))',
             fmt=MONEY, font=W.BOLD, fill=PatternFill("solid", fgColor=B.T_NEUTRAL))
        cell(ws, r, 8, f'=G{r}*{t[0]["mgr_pct"]}', fmt=MONEY2,
             fill=PatternFill("solid", fgColor=B.T_NEUTRAL))
        cell(ws, r, 9, f'=IFERROR(G{r}*{t[0]["w_pct"]}/{max(s["staff"] - s["mgr"], 1)},0)',
             fmt=MONEY2, font=W.BOLD, fill=PatternFill("solid", fgColor=B.T_NEUTRAL))
        ws.row_dimensions[r].height = 20
        r += 1

    r = W.band(ws, r + 1, NC, "③ شرائح الحافز كما هي معتمدة — ولكل محطة جدولها")
    W.header(ws, r, ["#", "المحطة", "الشريحة", "الوعاء", "نسبة المدير", "مبلغ المدير",
                     "نسبة العمال", "مبلغ العمال", "للعامل الواحد"])
    r += 1
    for i, s in enumerate(S, 1):
        nw = max(s["staff"] - s["mgr"], 1)
        for j, t in enumerate(s["tiers"]):
            band = {0: "٨٥٪ إلى ٩٩٪", 1: "١٠٠٪ بالضبط", 2: "١٠٠٪ فأكثر"}[j]
            cell(ws, r, 1, i if j == 0 else "")
            cell(ws, r, 2, s["name"] if j == 0 else "", font=W.BOLD, align=W.RGT)
            cell(ws, r, 3, band, font=W.BOLD)
            cell(ws, r, 4, t["pot"], fmt=MONEY, font=W.BOLD)
            cell(ws, r, 5, t["mgr_pct"], fmt="0%")
            cell(ws, r, 6, f'=D{r}*E{r}', fmt=MONEY2,
                 fill=PatternFill("solid", fgColor=B.T_NEUTRAL))
            cell(ws, r, 7, t["w_pct"], fmt="0%")
            cell(ws, r, 8, f'=D{r}*G{r}', fmt=MONEY2,
                 fill=PatternFill("solid", fgColor=B.T_NEUTRAL))
            cell(ws, r, 9, f'=IFERROR(H{r}/{nw},0)', fmt=MONEY2, font=W.BOLD,
                 fill=PatternFill("solid", fgColor=B.T_NEUTRAL))
            r += 1
    ws.cell(r + 1, 2, "ملاحظة تستحق قراراً: نسبة المدير ٣٠٪ في الشرائع و١٥٪ في العزيزية — "
                      "المعيار غير موحَّد بين المحطتين.").font = W.SMALL
    ws.cell(r + 1, 2).alignment = W.RGT
    return ws


# ═══════════════════════════════════════════════════════ ③ المبيعات اليومية
def daily(wb, S, idx=2):
    NC = 10
    ws = sheet(wb, "المبيعات اليومية", [4, 7, 16, 16, 16, 15, 15, 14, 14, 26],
               "المبيعات اليومية — سبتمبر ٢٠٢٦",
               "الفعلي مقابل ميزانية اليوم، والتراكمي مقابل ما يقابله — لا مقابل الشهر كله",
               NC, idx)
    r = 3
    for s in S:
        d = s["daily"]
        prods = [k for k in d if k in ("91", "95", "D")]
        r = W.band(ws, r, NC, f"{s['name']} · {s['code']} — "
                              f"ميزانية اليوم {arn(format(s['bu']/DAYS_IN_MONTH, ',.0f'))} ريال")
        W.header(ws, r, ["#", "اليوم"] + [f"بنزين {p}" if p != "D" else "ديزل" for p in prods]
                 + [""] * (3 - len(prods)) + ["إجمالي اليوم", "ميزانية اليوم",
                                              "الفارق", "التراكمي", "التحقيق التراكمي"])
        r += 1
        first = r
        daily_bu = s["bu"] / DAYS_IN_MONTH
        days = sorted({int(dd) for p in prods for dd in d[p]})
        for n, dd in enumerate(days, 1):
            cell(ws, r, 1, n)
            cell(ws, r, 2, dd)
            for j, p in enumerate(prods):
                cell(ws, r, 3 + j, d[p].get(str(dd)), fmt=MONEY2)
            for j in range(len(prods), 3):
                cell(ws, r, 3 + j, "")
            cols = [get_column_letter(3 + j) for j in range(len(prods))]
            cell(ws, r, 6, "=" + "+".join(f"{c}{r}" for c in cols), fmt=MONEY, font=W.BOLD,
                 fill=PatternFill("solid", fgColor=B.T_NEUTRAL))
            cell(ws, r, 7, daily_bu, fmt=MONEY)
            cell(ws, r, 8, f"=F{r}-G{r}", fmt=MONEY,
                 fill=PatternFill("solid", fgColor=B.T_NEUTRAL))
            cell(ws, r, 9, f"=SUM($F${first}:F{r})", fmt=MONEY, font=W.BOLD,
                 fill=PatternFill("solid", fgColor=B.T_NEUTRAL))
            cell(ws, r, 10, f"=IFERROR(I{r}/(G{r}*{n}),0)", fmt=PCT, font=W.BOLD,
                 fill=PatternFill("solid", fgColor=B.T_NEUTRAL))
            r += 1
        last = r - 1
        cell(ws, r, 2, "الإجمالي", font=W.BOLD, fill=PatternFill("solid", fgColor=B.T_BAND))
        for j in range(len(prods)):
            cl = get_column_letter(3 + j)
            cell(ws, r, 3 + j, f"=SUM({cl}{first}:{cl}{last})", fmt=MONEY, font=W.BOLD,
                 fill=PatternFill("solid", fgColor=B.T_BAND))
        for j in range(len(prods), 3):
            cell(ws, r, 3 + j, "", fill=PatternFill("solid", fgColor=B.T_BAND))
        for col, f in ((6, f"=SUM(F{first}:F{last})"), (7, f"=SUM(G{first}:G{last})"),
                       (8, f"=SUM(H{first}:H{last})"), (9, f"=I{last}"),
                       (10, f"=IFERROR(F{r}/G{r},0)")):
            cell(ws, r, col, f, fmt=PCT if col == 10 else MONEY, font=W.BOLD,
                 fill=PatternFill("solid", fgColor=B.T_BAND))
        cell(ws, r, 1, "", fill=PatternFill("solid", fgColor=B.T_BAND))
        ws.conditional_formatting.add(f"H{first}:H{last}", CellIsRule(
            operator="lessThan", formula=["0"],
            fill=PatternFill("solid", fgColor=B.T_BAD)))
        ws.conditional_formatting.add(f"H{first}:H{last}", CellIsRule(
            operator="greaterThanOrEqual", formula=["0"],
            fill=PatternFill("solid", fgColor=B.T_GOOD)))
        ws.conditional_formatting.add(f"F{first}:F{last}", DataBarRule(
            start_type="min", end_type="max", color=B.ORANGE))
        r += 3
    return ws


# ═══════════════════════════════════════════════════════ ④ المستهدف
def target(wb, S, idx=3):
    NC = 13
    ws = sheet(wb, "المستهدف", [4, 20, 9] + [12] * 8 + [12, 34],
               "من أين جاءت ميزانية سبتمبر",
               "الميزانية = مبيعات أغسطس × (١ + نسبة النمو) — وهذا يثبّتها على شهر واحد",
               NC, idx)
    r = W.band(ws, 3, NC, "① مبيعات ٢٠٢٦ شهراً بشهر — ومنها تُشتقّ الميزانية")
    months = [m for m, _ in S[0]["months"]]
    W.header(ws, r, ["#", "المحطة", "الكود"] + [MONTH_AR[m.split("-")[1]] for m in months]
             + ["ميزانية سبتمبر", "الطريقة"])
    r += 1
    for i, s in enumerate(S, 1):
        cell(ws, r, 1, i)
        cell(ws, r, 2, s["name"], font=W.BOLD, align=W.RGT)
        cell(ws, r, 3, s["code"])
        for j, (m, v) in enumerate(s["months"]):
            cell(ws, r, 4 + j, v, fmt=MONEY)
        last_col = get_column_letter(3 + len(months))
        cell(ws, r, 4 + len(months), f"={last_col}{r}*(1+{s['growth']})", fmt=MONEY,
             font=W.BOLD, fill=PatternFill("solid", fgColor=B.T_NEUTRAL))
        cell(ws, r, 5 + len(months), f"أغسطس × (١ + {arn(format(s['growth']*100, '.0f'))}٪)", align=W.WRAP)
        ws.row_dimensions[r].height = 20
        r += 1

    r = W.band(ws, r + 1, NC, "② وما يُفترض أن يُراجَع في هذه الطريقة")
    notes = [
        ("تثبيتها على شهر واحد", "أغسطس وحده يحدّد سبتمبر. لو كان أغسطس شاذّاً — "
                                 "موسماً أو انقطاعاً — انتقل شذوذه إلى المستهدف كله."),
        ("الشرائع مثال حيّ", "يناير ٥٦٨٬٥٦٦ ثم استقرار حول ٢٦٥٬٠٠٠ منذ فبراير. "
                             "الأساس المنخفض ليس أداءً بل أثر تغيّر مدخل المحطة."),
        ("بلا موسمية", "سبتمبر ليس أغسطس. المستهدف يحتاج معامل شهر من سنة سابقة، "
                       "وبيانات ٢٠٢٥ موجودة في الملف ولم تُستعمل."),
        ("نسبة نمو بلا سند", "٦٪ للشرائع و٧٪ للعزيزية — مصدرهما غير مذكور، "
                             "ويجب أن يُكتب إلى جانبهما."),
    ]
    W.header(ws, r, ["#", "الملاحظة"] + [""] * 9 + ["", "التفصيل"])
    r += 1
    for i, (t, d) in enumerate(notes, 1):
        cell(ws, r, 1, i)
        cell(ws, r, 2, t, font=W.BOLD, align=W.RGT)
        for c in range(3, 13): cell(ws, r, c, "")
        cell(ws, r, 13, d, align=W.WRAP)
        ws.row_dimensions[r].height = 32
        r += 1
    return ws


# ═══════════════════════════════════════════════════════ ⑤ بطاقة الأداء
def scorecard(wb, idx=4):
    NC = 9
    ws = sheet(wb, "بطاقة الأداء", [4, 26, 10, 18, 10, 12, 12, 10, 30],
               "بطاقة أداء المحطة — بالريال السعودي",
               "كانت وحدة القياس في الملف الأصلي JOD (ديناراً أردنياً) — قالبٌ لم يُحوَّل",
               NC, idx)
    KPI = [
        ("تقييم خرائط جوجل", "نقطة", 0.10, 90, 100, "الظهور والسمعة"),
        ("مبيعات الوقود", "م٣", 0.50, 265, 250, "المحرّك الأول"),
        ("جودة الخدمة", "٪", 0.15, 90, 90, "قياس ميداني"),
        ("التدقيق الخارجي", "٪", 0.10, 90, 90, "جهة مستقلة"),
        ("التدقيق الداخلي — النظافة", "٪", 0.05, 95, 95, "داخلي"),
        ("التدقيق الداخلي — السلوك", "٪", 0.05, 90, 90, "داخلي"),
    ]
    r = W.band(ws, 3, NC, "① المؤشرات وأوزانها")
    W.header(ws, r, ["#", "المؤشر", "الوحدة", "الفئة", "الوزن", "المستهدف",
                     "المحقَّق", "التحقيق", "المصدر"])
    r += 1
    first = r
    for i, (k, u, w, act, obj, src) in enumerate(KPI, 1):
        cell(ws, r, 1, i)
        cell(ws, r, 2, k, font=W.BOLD, align=W.RGT)
        cell(ws, r, 3, u)
        cell(ws, r, 4, "مبيعات" if "مبيعات" in k else "خدمة وتدقيق")
        cell(ws, r, 5, w, fmt="0%", fill=PatternFill("solid", fgColor=B.T_GOLD))
        cell(ws, r, 6, obj, fill=PatternFill("solid", fgColor=B.T_GOLD))
        cell(ws, r, 7, act, fmt="#,##0.00", fill=PatternFill("solid", fgColor=B.T_GOLD))
        cell(ws, r, 8, f"=IFERROR(G{r}/F{r},0)", fmt=PCT, font=W.BOLD,
             fill=PatternFill("solid", fgColor=B.T_NEUTRAL))
        cell(ws, r, 9, src, align=W.WRAP)
        r += 1
    last = r - 1
    cell(ws, r, 2, "الإجمالي المرجَّح", font=W.BOLD,
         fill=PatternFill("solid", fgColor=B.T_BAND), align=W.RGT)
    for c in (1, 3, 4, 6, 7, 9):
        cell(ws, r, c, "", fill=PatternFill("solid", fgColor=B.T_BAND))
    cell(ws, r, 5, f"=SUM(E{first}:E{last})", fmt="0%", font=W.BOLD,
         fill=PatternFill("solid", fgColor=B.T_BAND))
    cell(ws, r, 8, f"=SUMPRODUCT(E{first}:E{last},H{first}:H{last})/SUM(E{first}:E{last})",
         fmt=PCT, font=W.BOLD, fill=PatternFill("solid", fgColor=B.T_BAND))
    ws.conditional_formatting.add(f"H{first}:H{last}", CellIsRule(
        operator="greaterThanOrEqual", formula=["1"],
        fill=PatternFill("solid", fgColor=B.T_GOOD)))
    ws.conditional_formatting.add(f"H{first}:H{last}", CellIsRule(
        operator="lessThan", formula=["1"],
        fill=PatternFill("solid", fgColor=B.T_BAD)))

    r += 2
    r = W.band(ws, r, NC, "② الشروط النوعية — تُمنع بها الصرف ولا تُزاد بها")
    COND = ["متابعة يومية للمبيعات", "المتابعة الدورية على النظافة",
            "الحرص على تقديم أفضل خدمة", "تقديم مناديل مجاناً للعملاء",
            "تنظيف الزجاج الأمامي للسيارات", "تدريب العمال ومتابعتهم"]
    W.header(ws, r, ["#", "الشرط"] + [""] * 5 + ["مستوفى؟", "ملاحظة المشرف"])
    r += 1
    dv = DataValidation(type="list", formula1='"نعم,لا"', allow_blank=True)
    ws.add_data_validation(dv)
    for i, c in enumerate(COND, 1):
        cell(ws, r, 1, i)
        cell(ws, r, 2, c, font=W.BOLD, align=W.RGT)
        for k in range(3, 8): cell(ws, r, k, "")
        x = cell(ws, r, 8, "نعم", fill=PatternFill("solid", fgColor=B.T_GOLD), font=W.BOLD)
        dv.add(x)
        cell(ws, r, 9, "", fill=PatternFill("solid", fgColor=B.T_GOLD))
        r += 1
    ws.cell(r + 1, 2, "الشروط النوعية بوابة لا وزن: عدم استيفاء أيٍّ منها يمنع الصرف، "
                      "واستيفاؤها جميعاً لا يرفع المبلغ.").font = W.SMALL
    ws.cell(r + 1, 2).alignment = W.RGT
    return ws


# ═══════════════════════════════════════════════════════ ⑥ المنهجية
def method(wb, S, idx=5):
    NC = 6
    ws = sheet(wb, "المنهجية", [4, 30, 26, 26, 4, 46],
               "المنهجية — ما صُحِّح وما يلزم",
               "المصدر: ملفّا Stations_Staff_Incentive لـMK019 وMK072 · سبتمبر ٢٠٢٦",
               NC, idx)
    r = W.band(ws, 3, NC, "① ما صُحِّح عن الملف الأصلي")
    W.header(ws, r, ["#", "الموضع", "ما كان", "ما صار", "", "لماذا"])
    r += 1
    FIX = [
        ("احتساب التحقيق", "الفعلي ١٩ يوماً ÷ ميزانية ٣٠ يوماً = ٥٩٪",
         "÷ ميزانية ١٩ يوماً = ٩٢٫٥٪ و٩٣٫١٪",
         "قسمة جزءٍ على كلٍّ تُسقط الحافز عن محطتين مستحقّتين"),
        ("وحدة العملة", "JOD في بطاقة الأداء", "ريال سعودي",
         "قالب أردني نُقل بلا تحويل — والمبالغ تُصرف بالريال"),
        ("طبيعة الأرقام", "قيم مطبوعة لا تتغيّر", "صيغ حيّة تتحدّث بتغيّر المُدخَل",
         "الملف أداة تُشغَّل، لا صورة تُطبع"),
        ("عدد الملفات", "ملف لكل محطة", "ملف واحد يقبل أي عدد",
         "لا مقارنة بين محطتين في ملفين منفصلين"),
        ("حجم الورقة اليومية", "٢٬١٥٣ صفاً × ١٢٩ عموداً وفيها ٣٠ رقماً",
         "جدول بعدد أيام الشهر", "الفراغ يُبطئ الملف ويخفي الرقم"),
        ("ترتيب القراءة", "من اليسار", "من اليمين — عربي بالكامل",
         "الفريق يقرأ بالعربية"),
    ]
    for i, (a, b, c, d) in enumerate(FIX, 1):
        cell(ws, r, 1, i)
        cell(ws, r, 2, a, font=W.BOLD, align=W.RGT)
        cell(ws, r, 3, b, align=W.WRAP, fill=PatternFill("solid", fgColor=B.T_BAD))
        cell(ws, r, 4, c, align=W.WRAP, fill=PatternFill("solid", fgColor=B.T_GOOD))
        cell(ws, r, 5, "")
        cell(ws, r, 6, d, align=W.WRAP)
        ws.row_dimensions[r].height = 34
        r += 1

    r = W.band(ws, r + 1, NC, "② ما يلزم لإقرار الحافز — ولم يصل بعد")
    W.header(ws, r, ["#", "المطلوب", "من", "لماذا", "", "الأثر إن لم يصل"])
    r += 1
    NEED = [
        ("توحيد نسبة المدير", "الموارد البشرية", "٣٠٪ في الشرائع و١٥٪ في العزيزية",
         "عاملان في محطتين يأخذان المبلغ نفسه تقريباً بينما مديراهما يختلفان"),
        ("معامل موسمية شهرية", "العمليات", "مبيعات ٢٠٢٥ موجودة ولم تُستعمل",
         "المستهدف يبقى مثبَّتاً على شهر واحد"),
        ("سند نسبة النمو", "الإدارة التجارية", "٦٪ و٧٪ بلا مصدر مكتوب",
         "لا يمكن الدفاع عن المستهدف أمام من لم يحقّقه"),
        ("ربط بطاقة الأداء بالصرف", "الموارد البشرية",
         "البطاقة تُحسب ولا تؤثّر في المبلغ",
         "نقيس ستة مؤشرات ونصرف على واحد"),
    ]
    for i, (a, b, c, d) in enumerate(NEED, 1):
        cell(ws, r, 1, i)
        cell(ws, r, 2, a, font=W.BOLD, align=W.RGT)
        cell(ws, r, 3, b, align=W.WRAP)
        cell(ws, r, 4, c, align=W.WRAP)
        cell(ws, r, 5, "")
        cell(ws, r, 6, d, align=W.WRAP, fill=PatternFill("solid", fgColor=B.T_GOLD))
        ws.row_dimensions[r].height = 34
        r += 1
    return ws


def build():
    S = list(json.load(open(SRC, encoding="utf-8")).values())
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    dashboard(wb, S, 0)
    calc(wb, S, 1)
    daily(wb, S, 2)
    target(wb, S, 3)
    scorecard(wb, 4)
    method(wb, S, 5)
    wb.save(OUT)
    return OUT, [ws.title for ws in wb.worksheets]


if __name__ == "__main__":
    f, names = build()
    print("saved:", f, "|", " · ".join(names))
