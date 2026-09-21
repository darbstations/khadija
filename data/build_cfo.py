# -*- coding: utf-8 -*-
"""تقرير المالية يوليو ٢٠٢٦ في إكسل عربي — DARB_Commercial_Report_Jul2026

   الأرقام منقولة عن النسخة المصححة (١٣ صفحة) بآلاف الريالات، وكل خلية
   نتيجةٍ صيغةٌ حيّة لا قيمة مطبوعة، فما يُطابَق يُرى مطابَقاً.

   يُشغَّل من جذر المستودع:  PYTHONPATH=data python3 data/build_cfo.py
"""
import csv, json
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule, DataBarRule
from openpyxl.chart import LineChart, BarChart, Reference

import brand as B
import worker_model as W

OUT = "docs/تقرير-المالية-يوليو-٢٠٢٦.xlsx"
CFO_ST = "outlets/data/cfo-stations.csv"

MONEY = '#,##0;[Red](#,##0)'
PCT = '0.0%'
PCT2 = '0.00%'
NUM2 = '#,##0.00'

# ── ص٥ و٦: قائمة الدخل — يوليو مفرداً والتراكمي · بآلاف الريالات
#    [الفعلي، الموازنة، العام الماضي]
PL_JUL = [
    ("إيرادات الوقود",            195115, 209322,  75048, "fuel"),
    ("تكلفة المبيعات",           -188732, -202095, -70881, "fuel"),
    ("هامش المساهمة — الوقود",      6382,   7228,   4167, "sub"),
    ("الإيراد العقاري",             5495,   5964,   3808, "rent"),
    ("خصومات الإيجار",              -516,   -537,    -56, "rent"),
    ("صافي الإيراد العقاري",        4979,   5427,   3753, "sub"),
    ("إهلاك أصول حق الاستخدام",    -2393,  -2582,  -1953, "rent"),
    ("تكلفة التمويل",              -2443,  -2394,  -1454, "rent"),
    ("هامش المساهمة — العقارات",     144,    451,    345, "sub"),
    ("مصاريف التشغيل",             -4950,  -4330,  -2549, "op"),
    ("مجمل الربح",                  1576,   3349,   1963, "sub"),
    ("بيع وتسويق + عمومية وإدارية", -2156,  -1811,  -2286, "op"),
    ("الربح قبل الفوائد والزكاة",   -580,   1538,   -322, "sub"),
    ("فوائد القروض",                -464,   -365,   -101, "op"),
    ("الزكاة",                        -40,    -29,      0, "op"),
    ("صافي الدخل",                 -1084,   1143,   -423, "net"),
]
PL_YTD = [
    ("إيرادات الوقود",           1173578, 1219570, 411451, "fuel"),
    ("تكلفة المبيعات",          -1133905, -1177404, -391474, "fuel"),
    ("هامش المساهمة — الوقود",     39673,  42166,  19976, "sub"),
    ("الإيراد العقاري",            34433,  36416,  27763, "rent"),
    ("خصومات الإيجار",             -2968,  -3277,  -2876, "rent"),
    ("صافي الإيراد العقاري",       31465,  33138,  24887, "sub"),
    ("إهلاك أصول حق الاستخدام",   -15347, -15999, -12391, "rent"),
    ("تكلفة التمويل",             -15525, -14569,  -9828, "rent"),
    ("هامش المساهمة — العقارات",     593,   2570,   2669, "sub"),
    ("مصاريف التشغيل",            -28704, -27760, -13962, "op"),
    ("مجمل الربح",                 11563,  16976,   8683, "sub"),
    ("بيع وتسويق + عمومية وإدارية", -13958, -10927, -12063, "op"),
    ("الربح قبل الفوائد والزكاة",  -2396,   6048,  -3380, "sub"),
    ("إيرادات أخرى",                  24,      0,      0, "op"),
    ("فوائد القروض",               -3043,  -1653,   -459, "op"),
    ("الزكاة",                      -120,   -110,   -101, "op"),
    ("صافي الدخل",                 -5535,   4285,  -3941, "net"),
]

# ── ص٧: الأداء الشهري (أربعة أشهر معروضة) والصافي الشهري لسبعة أشهر
MONTHS4 = ["مارس", "مايو", "يونيو", "يوليو"]
MONTHLY = [
    ("إجمالي الإيرادات",   [158383, 170580, 179993, 200094], 1205042),
    ("تكلفة المبيعات",     [-150013, -160193, -169666, -188732], -1133905),
    ("الدخل الإيجاري",     [4748, 4908, 4748, 4979], 31465),
    ("مصاريف التشغيل",     [-3602, -4516, -4149, -4950], -28704),
    ("مجمل الربح",         [1931, 892, 1418, 1576], 11563),
    ("عمومية وإدارية",     [-1667, -1846, -1667, -1859], -11526),
    ("بيع وتسويق",         [-317, -513, -317, -297], -2432),
    ("فوائد القروض",       [-492, -541, -492, -464], -3043),
    ("صافي الدخل",         [-370, -2048, -1082, -1084], -5535),
]
MONTHS7 = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو"]
NET7 = [293, -414, -370, -832, -2048, -1082, -1084]

# ── ص٨: قائمة دخل كاملة حسب نموذج العمل
MODELS = ["استثماري", "إيجاري", "تشغيلي", "امتياز"]
BYMODEL = {
    "stations": [6, 17, 26, 92],
    "litres_m": [91.8, 77.7, 141.2, 350.8],
    "revenue":  [174309, 146564, 268915, 583790],
    "cogs":     [-163544, -137777, -252931, -579654],
    "fuel_cm":  [10765, 8787, 15984, 4136],
    "rent_inc": [19050, 10834, 2512, 2037],
    "rou_dep":  [-5635, -8874, -557, -281],
    "finance":  [-5032, -8165, -631, -1696],
    "rent_cm":  [6427, -7141, 1248, 60],
    "opex":     [-10167, -4378, -13896, -144],
    "gross":    [7025, -2734, 3295, 4052],
    "selling":  [-597, -597, -717, -478],
    "admin":    [-2921, -2921, -3505, -2337],
    "interest": [-761, -761, -913, -609],
    "net":      [2722, -7038, -1869, 609],
}

# ── ص٤: شبكة المحطات — أربع موازنات ثم الفعلي والعام الماضي وموازنة ديسمبر
#    مستخرجة بمجاميع تحقُّق تنطبق كلها: ١٧٦/١٨٠/١٩١/١٩٨/١٤٦/٦٨/٢٣٢
NET_COLS = ["موازنة أبريل", "موازنة مايو", "موازنة يونيو", "موازنة يوليو",
            "يوليو فعلي", "يوليو ٢٠٢٥", "موازنة ديسمبر"]
NETWORK = {"استثماري": [6, 6, 6, 6, 6, 5, 8],
           "إيجاري":   [20, 20, 24, 24, 17, 13, 28],
           "تشغيلي":   [35, 35, 35, 38, 29, 8, 49],
           "امتياز":   [115, 119, 126, 130, 94, 42, 147]}

# ── ص٢ و١٢: ما صُحِّح · وص١٣: ما بقي مفتوحاً
FIXES = [
    ("توحيد عدد المحطات", "١٤٦ (جدول الشبكة) و١٤١ (النماذج) و٢٢١ (عناوين النِّعم)",
     "١٤٦ للمواقع النشطة، والتعريفات مكتوبة"),
    ("تسمية رابحة وخاسرة", "بعض المحطات مصنَّفة عكس صافي دخلها",
     "التسمية الآن من صافي دخل المحطة نفسها"),
    ("أعلى عشر قوائم", "المجموع المعلن ١٢٬٨٢٠ وصفوفه تجمع ١٢٬٦٧٨",
     "أُعيد حسابها فصار المجموع يساوي صفوفه"),
    ("النسب المشوَّهة", "هوامش ٤٩٢٫٧٧٪ و١٢٧٫٣٣٪ لمحطات بلا مليون ريال إيراد",
     "استُبعدت من الترتيب"),
    ("المحطة المكررة", "MK289 تظهر في قائمتي الربح والخسارة بلا تمييز",
     "مرة واحدة لكل نموذج مع وسم"),
    ("أرقام السرد", "«+١٠٧ عن العام الماضي» و«٣٫٠M محاور»",
     "+٧٨ و٢٫٩M · والكميات ٦٦١٫٤ مليون لتر تُذكر مرة واحدة"),
    ("سلامة الرسوم", "رسمٌ يحمل ثلاث سلاسل غير متّسقة", "حُذف بدل إعادة إنتاجه"),
]
GAPS = [
    ("فجوة مجمل الربح", "قوائم المحطات تجمع +١٨٬٣٧٥ قبل التحميل، وجدول النماذج ١١٬٦٣٧",
     "٦٬٧٣٨ ألف ريال", "بنود التكلفة المُدرجة على مستوى المحطة مقابل مستوى النموذج"),
    ("عدد المحطات بلا مبيعات", "عناوين النِّعم ١ و١١ و١٢ و٨ · وصفحات السرد ٤ و١٤ و٢٣ و١١",
     "بلا رقم واحد", "«+٦١ محطة امتياز» أي ٦٨ لم تُصحَّح"),
    ("سجل المواقع المعتمَد", "٣١ موقعاً بلا مبيعات · و١٨ محطة خاسرة قابلة للتحقق",
     "ادعاء «٦٦ محطة خاسرة» لم يُحسم", "يحتاج سجلاً معتمَداً بتاريخ يوليو ٢٠٢٦"),
    ("إيراد ٢٠٢٥ والموازنة", "الشريحة الوحيدة الحاملة للربع أظهرت ثلاث سلاسل غير متّسقة",
     "غير متاح حسب النموذج", "مطلوب إيراد ٢٠٢٥ وموازنته لكل نموذج"),
    ("مجمَّع الملحق", "JA070 وJA234 وMK310 لا تظهر في أي نموذج", "ثلاث محطات",
     "ملحق المحطات يربّي مواقع خارج التجميع"),
]

CPL_NOTE = ("التكلفة والقيمة لكل لتر تُشتقّ من هذه القائمة وحدها — "
            "لا من تقرير ربحية المحطات، فمصاريفه التشغيلية نصف ما هنا")


def sheet(wb, name, widths, title, sub, nc, idx=None):
    ws = wb.create_sheet(name) if idx is None else wb.create_sheet(name, idx)
    W.setup(ws, widths, freeze="A4")
    W.title(ws, title, sub, nc)
    return ws


def cell(ws, r, c, v, *, fmt=None, font=None, fill=None, align=None):
    x = ws.cell(r, c, v)
    if fmt: x.number_format = fmt
    x.font = font or W.BLACK
    if fill: x.fill = fill
    x.alignment = align or W.CTR
    x.border = W.BOX
    return x


def kind_fill(k):
    return {"sub": PatternFill("solid", fgColor=B.T_NEUTRAL),
            "net": PatternFill("solid", fgColor=B.T_BAND)}.get(k)


# ═══════════════════════════════════════════════════════ ① الملخص
def summary(wb, idx=0):
    NC = 7
    ws = sheet(wb, "الملخص", [4, 30, 16, 16, 16, 4, 40],
               "تقرير الأداء — يوليو ٢٠٢٦",
               "شركة درب للمحطات · يناير–يوليو ٢٠٢٦ · بآلاف الريالات · "
               "النسخة المصححة من مصدر ٢٤٠ صفحة", NC, idx)
    r = W.band(ws, 3, NC, "① المؤشرات الرئيسية — تراكمي سبعة أشهر")
    W.header(ws, r, ["#", "المؤشر", "الفعلي", "الموازنة", "العام الماضي", "", "القراءة"])
    r += 1
    KPI = [
        ("الإيرادات التراكمية", 1173578, 1219570, 411451, "+١٨٥٪ سنوياً · و−٣٫٨٪ عن الموازنة"),
        ("مجمل الربح", 11563, 16976, 8683, "+٣٣٪ سنوياً · و−٣١٫٩٪ عن الموازنة"),
        ("صافي الدخل", -5535, 4285, -3941, "الفارق عن الموازنة ٩٬٨٢٠ ألف ريال"),
        ("الكميات (مليون لتر)", 661.4, 693.3, 220.5, "×٣٫٠ عن العام الماضي · −٤٫٦٪ عن الموازنة"),
        ("المحطات النشطة", 146, 94, 68, "+٧٨ سنوياً · و+٥٢ عن خطة يوليو"),
    ]
    first = r
    for i, (k, a, b, c, note) in enumerate(KPI, 1):
        fmt = NUM2 if "مليون" in k else MONEY
        cell(ws, r, 1, i)
        cell(ws, r, 2, k, font=W.BOLD, align=W.RGT)
        cell(ws, r, 3, a, fmt=fmt, font=W.BOLD)
        cell(ws, r, 4, b, fmt=fmt)
        cell(ws, r, 5, c, fmt=fmt)
        cell(ws, r, 6, "")
        cell(ws, r, 7, note, align=W.WRAP)
        ws.row_dimensions[r].height = 20
        r += 1
    ws.conditional_formatting.add(f"C{first}:C{r-1}", CellIsRule(
        operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor=B.T_BAD)))

    r = W.band(ws, r + 1, NC, "② صافي الدخل شهراً بشهر — سالبٌ في كل شهر منذ فبراير")
    W.header(ws, r, ["#", "الشهر", "صافي الدخل", "", "", "", "ملاحظة"])
    r += 1
    mfirst = r
    for i, (m, v) in enumerate(zip(MONTHS7, NET7), 1):
        cell(ws, r, 1, i)
        cell(ws, r, 2, m, font=W.BOLD, align=W.RGT)
        cell(ws, r, 3, v, fmt=MONEY, font=W.BOLD,
             fill=PatternFill("solid", fgColor=B.T_GOOD if v > 0 else B.T_BAD))
        for k in (4, 5, 6): cell(ws, r, k, "")
        cell(ws, r, 7, "الشهر الوحيد الموجب" if v > 0 else "", align=W.WRAP)
        r += 1
    cell(ws, r, 2, "التراكمي", font=W.BOLD, fill=PatternFill("solid", fgColor=B.T_BAND), align=W.RGT)
    cell(ws, r, 3, f"=SUM(C{mfirst}:C{r-1})", fmt=MONEY, font=W.BOLD,
         fill=PatternFill("solid", fgColor=B.T_BAND))
    for k in (1, 4, 5, 6, 7): cell(ws, r, k, "", fill=PatternFill("solid", fgColor=B.T_BAND))

    ch = LineChart(); ch.title = "صافي الدخل شهرياً — ألف ريال"; ch.height = 7; ch.width = 16
    ch.add_data(Reference(ws, min_col=3, min_row=mfirst, max_row=mfirst + 6), titles_from_data=False)
    ch.set_categories(Reference(ws, min_col=2, min_row=mfirst, max_row=mfirst + 6))
    ch.series[0].graphicalProperties.line.solidFill = B.BAD
    ch.series[0].graphicalProperties.line.width = 28000
    ws.add_chart(ch, f"B{r + 2}")
    return ws


# ═══════════════════════════════════════════════════════ ② قائمة الدخل
def income(wb, idx=1):
    NC = 9
    ws = sheet(wb, "قائمة الدخل", [4, 30, 15, 15, 13, 13, 15, 13, 13],
               "قائمة الدخل — يوليو مفرداً والتراكمي",
               "بآلاف الريالات · والفروق محسوبة لا منقولة", NC, idx)
    r = 3
    for label, data in (("② تراكمي يوليو — سبعة أشهر", PL_YTD),
                        ("① يوليو مفرداً", PL_JUL)):
        r = W.band(ws, r, NC, label)
        W.header(ws, r, ["#", "البند", "الفعلي", "الموازنة", "الفرق", "الفرق ٪",
                         "العام الماضي", "الفرق", "الفرق ٪"])
        r += 1
        for i, (k, a, b, c, kind) in enumerate(data, 1):
            f = kind_fill(kind)
            cell(ws, r, 1, i, fill=f)
            cell(ws, r, 2, k, font=W.BOLD if kind in ("sub", "net") else W.BLACK,
                 align=W.RGT, fill=f)
            cell(ws, r, 3, a, fmt=MONEY, font=W.BOLD, fill=f)
            cell(ws, r, 4, b, fmt=MONEY, fill=f)
            cell(ws, r, 5, f"=C{r}-D{r}", fmt=MONEY, fill=f)
            cell(ws, r, 6, f'=IFERROR(E{r}/ABS(D{r}),"")', fmt=PCT, fill=f)
            cell(ws, r, 7, c, fmt=MONEY, fill=f)
            cell(ws, r, 8, f"=C{r}-G{r}", fmt=MONEY, fill=f)
            cell(ws, r, 9, f'=IFERROR(H{r}/ABS(G{r}),"")', fmt=PCT, fill=f)
            r += 1
        ws.conditional_formatting.add(f"E{r-len(data)}:E{r-1}", CellIsRule(
            operator="lessThan", formula=["0"],
            font=Font(name=B.FONT, size=10, color=B.BAD)))
        r += 2
    ws.cell(r, 2, CPL_NOTE).font = W.SMALL
    ws.cell(r, 2).alignment = W.RGT
    return ws


# ═══════════════════════════════════════════════════════ ③ الأداء الشهري
def monthly(wb, idx=2):
    NC = 7
    ws = sheet(wb, "الأداء الشهري", [4, 30, 14, 14, 14, 14, 16],
               "الأداء الشهري — أربعة أشهر معروضة",
               "الإيراد سجّل رقماً قياسياً في يوليو، والصافي أسوأ — فالمصاريف تنمو أسرع",
               NC, idx)
    r = W.band(ws, 3, NC, "الأشهر كما عرضها التقرير · بآلاف الريالات")
    W.header(ws, r, ["#", "البند"] + MONTHS4 + ["تراكمي يوليو"])
    r += 1
    rev_row = None
    for i, (k, vals, tot) in enumerate(MONTHLY, 1):
        f = kind_fill("net") if k == "صافي الدخل" else (
            kind_fill("sub") if k == "مجمل الربح" else None)
        cell(ws, r, 1, i, fill=f)
        cell(ws, r, 2, k, font=W.BOLD, align=W.RGT, fill=f)
        for j, v in enumerate(vals):
            cell(ws, r, 3 + j, v, fmt=MONEY, fill=f,
                 font=W.BOLD if k in ("صافي الدخل", "مجمل الربح") else W.BLACK)
        cell(ws, r, 7, tot, fmt=MONEY, font=W.BOLD,
             fill=f or PatternFill("solid", fgColor=B.T_NEUTRAL))
        if k == "إجمالي الإيرادات": rev_row = r
        if k == "صافي الدخل":
            ws.conditional_formatting.add(f"C{r}:F{r}", CellIsRule(
                operator="lessThan", formula=["0"],
                fill=PatternFill("solid", fgColor=B.T_BAD)))
        r += 1
    ch = BarChart(); ch.title = "الإيراد الشهري — ألف ريال"; ch.height = 7; ch.width = 14
    ch.add_data(Reference(ws, min_col=3, max_col=6, min_row=rev_row), titles_from_data=False)
    ch.set_categories(Reference(ws, min_col=3, max_col=6, min_row=r - len(MONTHLY) - 1))
    ws.add_chart(ch, f"B{r + 2}")
    return ws


# ═══════════════════════════════════════════════════════ ④ نماذج العمل
def bymodel(wb, idx=3):
    NC = 7
    ws = sheet(wb, "نماذج العمل", [4, 28, 15, 15, 15, 15, 15],
               "قائمة الدخل حسب نموذج العمل — تراكمي يوليو",
               "الاستثماري وحده رابح بعد التحميلات · والإيجاري يحمل الخسارة", NC, idx)
    r = W.band(ws, 3, NC, "① القائمة الكاملة · بآلاف الريالات")
    W.header(ws, r, ["#", "البند"] + MODELS + ["الإجمالي"])
    r += 1
    ROWS = [("عدد المحطات", "stations", None), ("الكميات (مليون لتر)", "litres_m", NUM2),
            ("الإيراد", "revenue", MONEY), ("تكلفة المبيعات", "cogs", MONEY),
            ("هامش المساهمة — الوقود", "fuel_cm", MONEY),
            ("الدخل الإيجاري", "rent_inc", MONEY),
            ("إهلاك حق الاستخدام", "rou_dep", MONEY),
            ("تكلفة التمويل", "finance", MONEY),
            ("هامش المساهمة — الإيجار", "rent_cm", MONEY),
            ("مصاريف تشغيل المحطات", "opex", MONEY),
            ("مجمل الربح", "gross", MONEY),
            ("بيع وتسويق (محمّل)", "selling", MONEY),
            ("عمومية وإدارية (محمّل)", "admin", MONEY),
            ("فوائد القروض (محمّل)", "interest", MONEY),
            ("صافي الربح", "net", MONEY)]
    first = r
    key_row = {}
    for i, (lbl, key, fmt) in enumerate(ROWS, 1):
        sub = key in ("fuel_cm", "rent_cm", "gross", "net")
        f = kind_fill("net") if key == "net" else (kind_fill("sub") if sub else None)
        cell(ws, r, 1, i, fill=f)
        cell(ws, r, 2, lbl, font=W.BOLD if sub else W.BLACK, align=W.RGT, fill=f)
        for j, v in enumerate(BYMODEL[key]):
            cell(ws, r, 3 + j, v, fmt=fmt or MONEY, font=W.BOLD if sub else W.BLACK, fill=f)
        cell(ws, r, 7, f"=SUM(C{r}:F{r})", fmt=fmt or MONEY, font=W.BOLD,
             fill=f or PatternFill("solid", fgColor=B.T_NEUTRAL))
        key_row[key] = r
        r += 1
    ws.conditional_formatting.add(f"C{key_row['net']}:G{key_row['net']}", CellIsRule(
        operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor=B.T_BAD),
        font=Font(name=B.FONT, size=10, bold=True, color=B.BAD)))

    r = W.band(ws, r + 1, NC, "② وبالهللة لكل لتر — وهنا يُقرأ الفارق")
    W.header(ws, r, ["#", "البند"] + MODELS + ["الشبكة"])
    r += 1
    CPL = [("سعر البيع المحقَّق", "revenue"), ("شراء الوقود", "cogs"),
           ("هامش المساهمة", "fuel_cm"), ("مصاريف تشغيل المحطة", "opex"),
           ("محمّلات المركز", None), ("صافي الربح", "net")]
    lit = key_row["litres_m"]
    for i, (lbl, key) in enumerate(CPL, 1):
        sub = key in ("fuel_cm", "net")
        f = kind_fill("net") if key == "net" else (kind_fill("sub") if sub else None)
        cell(ws, r, 1, i, fill=f)
        cell(ws, r, 2, lbl, font=W.BOLD if sub else W.BLACK, align=W.RGT, fill=f)
        for j in range(5):
            col = get_column_letter(3 + j)
            src = (f"{col}{key_row[key]}" if key else
                   f"({col}{key_row['selling']}+{col}{key_row['admin']}+{col}{key_row['interest']})")
            # التكاليف مخزَّنة سالبةً فتُعرض بمقدارها؛ والصافي يحتفظ بإشارته
            wrap = src if key == "net" else f"ABS({src})"
            cell(ws, r, 3 + j, f"={wrap}/({col}{lit}*1000)*100", fmt='+0.00;-0.00',
                 font=W.BOLD if sub else W.BLACK, fill=f)
        r += 1
    ws.cell(r + 1, 2, "التكاليف تُعرض بمقدارها، والصافي بإشارته — فالإيجاري "
                      "−٩٫٠٦ هللة لكل لتر والاستثماري +٢٫٩٧.").font = W.SMALL
    ws.cell(r + 1, 2).alignment = W.RGT
    return ws


# ═══════════════════════════════════════════════════════ ⑤ المحطات
def stations(wb, idx=4):
    NC = 8
    ws = sheet(wb, "المحطات", [4, 10, 26, 13, 14, 14, 12, 14],
               "أعلى عشر محطات ربحاً · وثمانية عشرة خاسرة",
               "هامش الوقود لا يفرّق بينهما — فالخسارة ليست في الوقود", NC, idx)
    R = list(csv.DictReader(open(CFO_ST, encoding="utf-8")))
    r = 3
    for kind, ttl in (("رابحة", "① أعلى عشر محطات ربحاً"),
                      ("خاسرة", "② المحطات الخاسرة — ثمانية عشرة")):
        g = [x for x in R if x["kind"] == kind]
        g.sort(key=lambda x: -float(x["net_k"]))
        r = W.band(ws, r, NC, ttl)
        W.header(ws, r, ["#", "الكود", "المحطة", "النموذج", "الإيراد التراكمي",
                         "صافي الدخل", "النسبة", "هامش الوقود"])
        r += 1
        first = r
        for i, x in enumerate(g, 1):
            net = float(x["net_k"])
            cell(ws, r, 1, i)
            cell(ws, r, 2, x["code"])
            cell(ws, r, 3, x["name"], font=W.BOLD, align=W.RGT)
            cell(ws, r, 4, x["model"])
            cell(ws, r, 5, float(x["rev_k"]), fmt=MONEY)
            cell(ws, r, 6, net, fmt=MONEY, font=W.BOLD)
            cell(ws, r, 7, f"=IFERROR(F{r}/E{r},0)", fmt=PCT2,
                 fill=PatternFill("solid", fgColor=B.T_NEUTRAL))
            cell(ws, r, 8, float(x["fuel_margin"]), fmt=PCT2, font=W.BOLD)
            r += 1
        last = r - 1
        cell(ws, r, 3, "المجموع", font=W.BOLD,
             fill=PatternFill("solid", fgColor=B.T_BAND), align=W.RGT)
        for c_, f_ in ((5, f"=SUM(E{first}:E{last})"), (6, f"=SUM(F{first}:F{last})"),
                       (8, f"=MEDIAN(H{first}:H{last})")):
            cell(ws, r, c_, f_, fmt=PCT2 if c_ == 8 else MONEY, font=W.BOLD,
                 fill=PatternFill("solid", fgColor=B.T_BAND))
        for c_ in (1, 2, 4, 7):
            cell(ws, r, c_, "", fill=PatternFill("solid", fgColor=B.T_BAND))
        ws.conditional_formatting.add(f"H{first}:H{last}", DataBarRule(
            start_type="num", start_value=0, end_type="num", end_value=0.08, color=B.ORANGE))
        ws.conditional_formatting.add(f"F{first}:F{last}", CellIsRule(
            operator="lessThan", formula=["0"],
            font=Font(name=B.FONT, size=10, bold=True, color=B.BAD)))
        r += 3
    ws.cell(r, 2, "وسيط هامش الخاسرة أعلى من الرابحة — والخسارة في عقد الإيجار "
                  "لا في هامش الوقود. انظر ورقة «نماذج العمل».").font = W.SMALL
    ws.cell(r, 2).alignment = W.RGT
    return ws


# ═══════════════════════════════════════════════════════ ⑥ الشبكة
def network(wb, idx=5):
    NC = 10
    ws = sheet(wb, "الشبكة", [4, 16] + [12] * 7 + [32],
               "شبكة المحطات — الفعلي مقابل الموازنة",
               "١٤٦ محطة فعلاً مقابل موازنة ١٩٨ · وهدف ديسمبر ٢٣٢ يحتاج ١٧ شهرياً "
               "والمتحقق ١١", NC, idx)
    r = W.band(ws, 3, NC, "المواقع النشطة حسب نموذج العمل")
    W.header(ws, r, ["#", "النموذج"] + NET_COLS + ["القراءة"])
    r += 1
    first = r
    NOTE = {"استثماري": "على الموازنة تماماً — وهو النموذج الرابح الوحيد",
            "إيجاري": "−٧ عن الموازنة · وهو الذي يحمل الخسارة",
            "تشغيلي": "−٩ عن الموازنة · و+٢١ عن العام الماضي",
            "امتياز": "−٣٦ عن الموازنة — ٦٩٪ من فجوة يوليو كلها منه"}
    for i, m in enumerate(MODELS, 1):
        v = NETWORK[m]
        cell(ws, r, 1, i)
        cell(ws, r, 2, m, font=W.BOLD, align=W.RGT)
        for j, x in enumerate(v):
            cell(ws, r, 3 + j, x, fmt="#,##0",
                 font=W.BOLD if j == 4 else W.BLACK,
                 fill=PatternFill("solid", fgColor=B.T_GOOD if j == 4 else
                                  (B.T_GOLD if j == 6 else B.T_NEUTRAL)) if j >= 4 else None)
        cell(ws, r, 10, NOTE[m], align=W.WRAP)
        ws.row_dimensions[r].height = 20
        r += 1
    cell(ws, r, 2, "الإجمالي", font=W.BOLD, fill=PatternFill("solid", fgColor=B.T_BAND),
         align=W.RGT)
    for j in range(7):
        col = get_column_letter(3 + j)
        cell(ws, r, 3 + j, f"=SUM({col}{first}:{col}{r-1})", fmt="#,##0", font=W.BOLD,
             fill=PatternFill("solid", fgColor=B.T_BAND))
    cell(ws, r, 1, "", fill=PatternFill("solid", fgColor=B.T_BAND))
    cell(ws, r, 10, "", fill=PatternFill("solid", fgColor=B.T_BAND))
    tot = r
    r += 2
    for lbl, f_, note in (
        ("الفرق عن موازنة يوليو", f"=G{tot}-F{tot}", "محطة دون الخطة — ٦٩٪ منها امتياز"),
        ("النمو عن العام الماضي", f"=G{tot}-H{tot}", "محطة · ×٢٫١ في اثني عشر شهراً"),
        ("الفجوة إلى ديسمبر", f"=I{tot}-G{tot}",
         "محطة خلال خمسة أشهر — أي ١٧ شهرياً مقابل ١١ متحققاً")):
        bad = "الفجوة" in lbl or "الفرق عن" in lbl
        fill = PatternFill("solid", fgColor=B.T_BAD if bad else B.T_GOOD)
        cell(ws, r, 2, lbl, font=W.BOLD, align=W.RGT, fill=fill)
        cell(ws, r, 3, f_, fmt="+#,##0;-#,##0", font=W.BOLD, fill=fill)
        cell(ws, r, 10, note, align=W.WRAP, fill=fill)
        for k in range(4, 10): cell(ws, r, k, "", fill=fill)
        cell(ws, r, 1, "", fill=fill)
        r += 1
    return ws


# ═══════════════════════════════════════════════════════ ⑦ التصحيحات
def integrity(wb, idx=6):
    NC = 6
    ws = sheet(wb, "التصحيحات والثغرات", [4, 26, 34, 30, 4, 36],
               "سلامة التقرير — ما صُحِّح وما بقي مفتوحاً",
               "هذه النسخة مصححة عن مصدر ٢٤٠ صفحة · ٣٤ خطأً صُحِّح وخمس ثغرات موثَّقة",
               NC, idx)
    r = W.band(ws, 3, NC, "① ما صُحِّح")
    W.header(ws, r, ["#", "الموضع", "ما كان في المصدر", "ما صار", "", ""])
    r += 1
    for i, (a, b, c) in enumerate(FIXES, 1):
        cell(ws, r, 1, i)
        cell(ws, r, 2, a, font=W.BOLD, align=W.RGT)
        cell(ws, r, 3, b, align=W.WRAP, fill=PatternFill("solid", fgColor=B.T_BAD))
        cell(ws, r, 4, c, align=W.WRAP, fill=PatternFill("solid", fgColor=B.T_GOOD))
        cell(ws, r, 5, ""); cell(ws, r, 6, "")
        ws.row_dimensions[r].height = 32
        r += 1
    r = W.band(ws, r + 1, NC, "② وما بقي مفتوحاً — لم يُقدَّر أيٌّ منها")
    W.header(ws, r, ["#", "الثغرة", "التفصيل", "المقدار", "", "المطلوب"])
    r += 1
    for i, (a, b, c, d) in enumerate(GAPS, 1):
        cell(ws, r, 1, i)
        cell(ws, r, 2, a, font=W.BOLD, align=W.RGT)
        cell(ws, r, 3, b, align=W.WRAP)
        cell(ws, r, 4, c, align=W.WRAP, font=W.BOLD,
             fill=PatternFill("solid", fgColor=B.T_GOLD))
        cell(ws, r, 5, "")
        cell(ws, r, 6, d, align=W.WRAP)
        ws.row_dimensions[r].height = 36
        r += 1
    return ws


def build():
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    summary(wb, 0); income(wb, 1); monthly(wb, 2); bymodel(wb, 3)
    stations(wb, 4); network(wb, 5); integrity(wb, 6)
    wb.save(OUT)
    return OUT, [w.title for w in wb.worksheets]


if __name__ == "__main__":
    f, names = build()
    print("saved:", f, "|", " · ".join(names))
