# -*- coding: utf-8 -*-
"""الحاسبة التجارية — ملف واحد تفاعلي لكل الأقسام، من كل بياناتنا لا من ملف واحد

   المصادر مجتمعةً: كاش إن للشبكة (٥٥ محطة) · قائمة الدخل يوليو ٢٠٢٦ ·
   تقرير ربحية الربع الثاني (٢٦ محطة) · التقرير اليومي لمحطة العمرة (٢٦٤ يوماً
   بالساعة) · مسح المنافسين · سجلّ الوحدات التجارية · سجلّ الشكاوى · ملفات
   العمّال والورديات والخزانات.

   كل خلية نتيجةٍ صيغةٌ حيّة لا رقمٌ مطبوع: من يغيّر مُدخَلاً يرى أثره فوراً،
   ومن يراجع يرى من أين جاء الرقم. والمُدخَلات وحدها بخلفية ذهبية.

   والترتيب ترتيب القاعدة: التشخيص أولاً، ثم الحاسبة، ثم المستهدف وكلفته،
   ثم ما لا نعرفه ومن يملك سدّه.

   يُشغَّل من جذر المستودع:  PYTHONPATH=data python3 data/build_calc.py
"""
import csv, json
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, DataBarRule
from openpyxl.chart import LineChart, BarChart, Reference

import brand as B
import worker_model as W

SRC = "data/mk007.json"
PLAN = "data/sales-plan.json"
COMP = "data/competitors.json"
LCOST = "data/litre-cost.json"
Q2COST = "outlets/data/q2-costs.csv"
OUT = "docs/الحاسبة-التجارية-درب.xlsx"

MONEY = '#,##0;[Red](#,##0)'
MONEY2 = '#,##0.00;[Red](#,##0.00)'
PCT = '0.0%'
PCT2 = '0.00%'
NUM = '#,##0'
NUM2 = '#,##0.00'
HAL = '0.00" هللة"'
LTR = '#,##0" لتر"'
TXN = '#,##0" معاملة"'
RIYAL4 = '0.0000" ريال"'

S_ST = "المحطات"
S_HR = "شكل الساعات"
S_CALC = "حاسبة المبيعات"
S_MK = "العمرة — بيان يومي"
S_H = "الساعة والوردية"
S_T = "خطة المستهدف"
S_C = "كلفة اللتر والمستهدف"

D = json.load(open(SRC, encoding="utf-8"))
Q = D["quarters"]; M = D["months"]; H = D["hourly"]; SH = D["shifts"]
T = D["target"]; C = D["cost"]; K = D["const"]; BL = D["base"]

P = json.load(open(PLAN, encoding="utf-8"))
X = json.load(open(COMP, encoding="utf-8"))
LC = json.load(open(LCOST, encoding="utf-8"))
ST = sorted(P["stations"], key=lambda s: -s["sar_total"])
SEGS = P["segments"]; NET = P["network"]; TOT = P["totals"]
CONTR = P["contracts"]; SG = P["shiftgap"]; LIT = P["litre"]

TBL = ""            # مرجع جدول البحث — يُملأ عند بناء ورقة المحطات
S_ST_RANGE = ""     # مدى أسماء المحطات للقائمة المنسدلة
HRS, HREF = {}, {}  # أسطر ورقة شكل الساعات · ووسيط كل تصنيف
DEFAULT_ST = "العمرة الجديدة"

# أسعار المضخة الفعلية من مبيعاتنا — شاملة الضريبة · للسياق لا للاشتقاق
PRICE_ = {"بنزين ٩١": 2.191, "بنزين ٩٥": 2.340, "ديزل": 1.796}

# شكل ساعات كل محطة من ملف الشبكة — معاملات كل ساعة عبر النافذة المقيسة
HOURS_RAW = {}
for _r in csv.DictReader(open("outlets/data/network-hours.csv", encoding="utf-8")):
    HOURS_RAW[_r["code"]] = [float(_r["h%02d" % _h] or 0) for _h in range(24)]

# عدد العمّال وحصة الوردية المسائية — تقرير العمّال والورديات
WORKERS = {}
for _r in csv.DictReader(open("outlets/data/worker-shifts.csv", encoding="utf-8")):
    try:
        WORKERS[_r["الرمز"]] = (float(_r["عدد العمّال"] or 0),
                                float(_r["حصة الوردية المسائية"] or 0))
    except ValueError:
        pass


# ═══════════════════════════════════════════════ أدوات الورقة
def sheet(wb, name, widths, title, sub, nc, freeze="A5"):
    ws = wb.create_sheet(name)
    W.setup(ws, widths, freeze=freeze)
    W.title(ws, title, sub, nc)
    W.rule(ws, 3, nc)
    return ws


def txt(ws, r, c, v, *, fmt=None, font=None, fill=None, align=None, box=True):
    x = ws.cell(r, c, v)
    x.font = font or W.BLACK
    x.alignment = align or W.CTR
    if fill: x.fill = fill
    if fmt: x.number_format = fmt
    if box: x.border = W.BOX
    return x


def merge(ws, r, c1, c2, v, *, font=None, fill=None, align=None, fmt=None):
    """خلية ممدودة: القيمة في أولى الأعمدة والتنسيق على كلها — فلا ينكسر الإطار"""
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    x = txt(ws, r, c1, v, font=font, fill=fill, align=align, fmt=fmt)
    for c in range(c1 + 1, c2 + 1):
        cell = ws.cell(r, c)
        cell.border = W.BOX
        if fill: cell.fill = fill
    return x


def note(ws, r, nc, text, fill=None, font=None, h=None):
    fl = fill or PatternFill("solid", fgColor=B.BG)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=nc)
    c = ws.cell(r, 1, text)
    c.font = font or Font(name=B.FONT, size=9, color=B.INK2)
    c.alignment = W.WRAP
    c.fill = fl
    c.border = W.BOX
    for j in range(2, nc + 1):
        cc = ws.cell(r, j); cc.border = W.BOX; cc.fill = fl
    ws.row_dimensions[r].height = h or 30
    return r + 1


def kpi(ws, r, c, label, value, fmt, tone=None, sub=None):
    fill = {"good": PatternFill("solid", fgColor=B.T_GOOD),
            "bad": PatternFill("solid", fgColor=B.T_BAD),
            "gold": PatternFill("solid", fgColor=B.T_GOLD)}.get(
                tone, PatternFill("solid", fgColor=B.T_NEUTRAL))
    col = {"good": B.D_GOOD, "bad": B.D_BAD, "gold": B.D_GOLD}.get(tone, B.INK)
    txt(ws, r, c, label, font=Font(name=B.FONT, size=9, color=B.INK2), fill=fill)
    x = txt(ws, r + 1, c, value, fmt=fmt, fill=fill,
            font=Font(name=B.FONT, size=13, bold=True, color=col))
    txt(ws, r + 2, c, sub or "", font=Font(name=B.FONT, size=8, color=B.INK3), fill=fill)
    return x


def sign_cf(ws, rng_, good_high=True):
    hi, lo = (B.T_GOOD, B.T_BAD) if good_high else (B.T_BAD, B.T_GOOD)
    dhi, dlo = (B.D_GOOD, B.D_BAD) if good_high else (B.D_BAD, B.D_GOOD)
    ws.conditional_formatting.add(rng_, CellIsRule(
        operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor=lo),
        font=Font(name=B.FONT, size=10, bold=True, color=dlo)))
    ws.conditional_formatting.add(rng_, CellIsRule(
        operator="greaterThan", formula=["0"], fill=PatternFill("solid", fgColor=hi),
        font=Font(name=B.FONT, size=10, bold=True, color=dhi)))


# ═══════════════════════════════════════════════ ① الملخص
def summary(wb):
    NC = 9
    ws = sheet(wb, "الملخص", [3, 30, 16, 16, 16, 16, 16, 4, 34],
               "القسم التجاري — الحاسبة والخطة في ملف واحد",
               f"{NET['stations']} محطة · {NET['volume']/1e6:,.1f} مليون لتر · "
               f"{NET['revenue']/1e6:,.1f} مليون ريال · {NET['visits']/1e6:,.2f} مليون زيارة · "
               f"{NET['raw_rows']:,} صفّاً من كاش إن", NC)
    r = 4
    r = note(ws, r, NC,
             "قاعدة هذا الملف: التشخيص قبل الوصفة. لا يدخل رقمٌ خانة «مبيعات» حتى "
             "يُستبعد قبله الوصولُ والبياناتُ والعقدُ والإمدادُ والرقابةُ والعمالة. "
             "ولذلك تجد في كل ورقة عموداً اسمه «البوابة» — محطةٌ عليها بوابة مفتوحة "
             "لا تدخل حملةً قبل إقفالها، لأن نتيجتها ستُنسب إلى الحملة وهي من البوابة.",
             fill=PatternFill("solid", fgColor=B.T_ORANGE),
             font=Font(name=B.FONT, size=10, bold=True, color=B.ORANGE), h=42)
    r += 1
    cards = [
        ("الفرصة المؤكَّدة", TOT["sar"], MONEY, "good", "ريال/سنة هامش مساهمة"),
        ("منها بعدد التعبئات", TOT["sar_txn"], MONEY, None,
         f"{TOT['gap_peak']:,.0f} معاملة/يوم في ساعات الذروة"),
        ("ومنها بحجم التعبئة", TOT["sar_fill"], MONEY, None,
         f"{TOT['upl_fill']:,.0f} لتر/يوم · تُفسَّر بالمزيج"),
        ("فجوة العمولة", abs(CONTR["gap_year"]), MONEY, "gold",
         "ريال/سنة — لا تحتاج حملةً بل تفاوضاً"),
        ("محطات دون المتوقَّع", TOT["below"], NUM, "bad",
         f"من {TOT['n']} محطة مقيسة"),
        ("وحدات تجارية شاغرة", P["units"]["vacant"], NUM, "bad",
         f"من {P['units']['total']} — إشغال "
         f"{P['units']['leased']/P['units']['total']*100:.0f}٪"),
    ]
    for i, (lab, v, f, tone, sub) in enumerate(cards):
        kpi(ws, r, 2 + i, lab, v, f, tone, sub)
    ws.row_dimensions[r + 1].height = 22
    r += 4

    # ── أين المال فعلاً
    r = W.band(ws, r, NC, "① أين المال فعلاً — وهل يحتاج بيعاً أصلاً؟")
    W.header(ws, r, ["", "الفرصة", "القيمة ر/سنة", "الصنف", "من يملكها", "",
                     "", "", "هل تحتاج حملةً؟"])
    r += 1
    m0 = r
    money = [
        ("فجوة العمولة في محطات التشغيل", abs(CONTR["gap_year"]), "عقد", "التجاري + المالية",
         "لا — تفاوض عقود. وهي أكبر رقم في الخطة كلها"),
        ("حجم التعبئة دون المتوقَّع", TOT["sar_fill"], "سلة", "التجاري",
         "جزئياً — وثلثا تباينها مزيج وقود لا سلوك، فتُفسَّر ولا تُطالَب"),
        ("فاقد ساعات الذروة", TOT["sar_txn"], "معاملات", "التجاري + العمليات",
         "لا — تغطية طاقم وسرعة خدمة. وهذا وحده ما يُطالَب به"),
        ("الوحدات الشاغرة المقيسة حركتها", None, "تأجير", "التأجير",
         f"لا — اكتتاب. {X['underwrite']['measured']} وحدة من "
         f"{X['underwrite']['total_vacant']} حركتها مقيسة اليوم"),
        ("كلفة آلية الهدية لو عُمِّمت", -C["gift_all_year"], "كلفة", "التجاري",
         "بالعكس — تُدفع على القائم والعائد من الإضافي · محطة العمرة وحدها"),
    ]
    for lab, v, kind, owner, need in money:
        txt(ws, r, 2, lab, align=W.RGT, font=W.BOLD)
        if v is None:
            txt(ws, r, 3, "غير مُسعَّرة", fill=PatternFill("solid", fgColor=B.T_GOLD),
                font=Font(name=B.FONT, size=9, bold=True, color=B.D_GOLD))
        else:
            txt(ws, r, 3, v, fmt=MONEY, fill=W.CALC, font=W.BOLD)
        txt(ws, r, 4, kind, font=Font(name=B.FONT, size=9, color=B.INK2))
        merge(ws, r, 5, 8, owner, align=W.RGT,
              font=Font(name=B.FONT, size=9, color=B.INK2))
        txt(ws, r, 9, need, align=W.WRAP, font=Font(name=B.FONT, size=9, color=B.INK2))
        ws.row_dimensions[r].height = 28
        r += 1
    sign_cf(ws, f"C{m0}:C{r-1}")
    r = note(ws, r, NC,
             "أكبر رقمٍ في الجدول لا يحتاج لتراً إضافياً ولا حملة — يحتاج جلسة تفاوض. "
             "وأضخم الفرص كثيراً ما تكون خارج دائرة البيع، وعلامة الخبرة أن تقول "
             "«هذه ليست مشكلة مبيعات» حين تكون كذلك.")
    r += 1

    # ── البوابات
    r = W.band(ws, r, NC, "② البوابات — حالاتٌ تُغلق قبل أي إنفاق")
    W.header(ws, r, ["", "الحالة المرصودة", "الفعل", "", "كيف", "", "",
                     "الصنف", "محطات عليها"])
    r += 1
    cond_n = {}
    for s in ST:
        for c_ in s["conds"]: cond_n[c_] = cond_n.get(c_, 0) + 1
    link = {"فجوة مطابقة تتجاوز ٥٪": "فجوة مطابقة",
            "تغيير مسار أو مدخل": "تغيير مسار",
            "وسيلة الدفع غير مسجّلة": "بيانات دفع مفقودة",
            "فاقد في ساعات الذروة": "فاقد ذروة",
            "تعبئة دون ما يفسّره مزيجها": "سلة دون المتوقَّع",
            "حصة الديزل ٣٠٪ فأكثر": "ديزل مرتفع",
            "عشر وحدات شاغرة فأكثر": "وحدات شاغرة"}
    for e in P["events"]:
        sign, act, how, kind, pri = e
        gate = pri == "hi"
        fill = PatternFill("solid", fgColor=B.T_GOLD) if gate else None
        txt(ws, r, 2, ("⛔ " if gate else "") + sign, align=W.RGT, fill=fill,
            font=Font(name=B.FONT, size=9, bold=gate,
                      color=B.D_GOLD if gate else B.INK))
        merge(ws, r, 3, 4, act, fill=fill,
              font=Font(name=B.FONT, size=9, bold=True,
                        color=B.D_GOLD if gate else B.ORANGE))
        merge(ws, r, 5, 7, how, align=W.RGT, fill=fill,
              font=Font(name=B.FONT, size=9, color=B.INK2))
        txt(ws, r, 8, kind, fill=fill, font=Font(name=B.FONT, size=9, color=B.INK2))
        n = cond_n.get(link.get(sign, ""), None)
        txt(ws, r, 9, n if n is not None else "—", fmt=NUM if n else None, fill=fill,
            font=Font(name=B.FONT, size=10, bold=True, color=B.D_BAD if n else B.INK3))
        ws.row_dimensions[r].height = 26
        r += 1
    r = note(ws, r, NC,
             "الأربع الأولى بوابات: تُغلق قبل أي حملة. والستّ الباقية أفعالٌ تجارية "
             "تبدأ بعدها. وعمود «محطات عليها» يقول كم محطةً ترفع كل علامة اليوم — "
             "والتفصيل في ورقة المحطات.",
             fill=PatternFill("solid", fgColor=B.T_GOLD),
             font=Font(name=B.FONT, size=9, bold=True, color=B.D_GOLD))
    r += 1

    # ── التصنيفات
    r = W.band(ws, r, NC, "③ التصنيف — مصطلحات الشركة نفسها من ملف التحليلات")
    W.header(ws, r, ["", "التصنيف", "محطات", "لتر/يوم", "لتر/زيارة", "ديزل٪",
                     "R²", "", "المحرّك الأول والفعل"])
    r += 1
    g0 = r
    for s in SEGS:
        txt(ws, r, 2, s["seg"], align=W.RGT, font=W.BOLD)
        txt(ws, r, 3, s["n"], fmt=NUM, fill=W.CALC)
        txt(ws, r, 4, s["lpd"], fmt=NUM, fill=W.CALC)
        txt(ws, r, 5, s["lpv"], fmt=NUM2, fill=W.CALC)
        txt(ws, r, 6, s["diesel"], fmt=PCT, fill=W.CALC)
        txt(ws, r, 7, s["r2"], fmt='0.00', fill=W.CALC)
        txt(ws, r, 8, "", box=False)
        txt(ws, r, 9, f"{s['driver']} — {s['action']}", align=W.WRAP,
            font=Font(name=B.FONT, size=9, color=B.INK2))
        ws.row_dimensions[r].height = 26
        r += 1
    txt(ws, r, 2, "الإجمالي", align=W.RGT, font=W.BOLD, fill=W.FS)
    for c_ in (3, 4):
        txt(ws, r, c_, f"=SUM({get_column_letter(c_)}{g0}:{get_column_letter(c_)}{r-1})",
            fmt=NUM, fill=W.FS, font=W.BOLD)
    for c_ in (5, 6, 7): txt(ws, r, c_, "", fill=W.FS)
    txt(ws, r, 8, "", box=False); txt(ws, r, 9, "", fill=W.FS)
    r += 1
    r = note(ws, r, NC,
             "R² هو مدى ما يفسّره التصنيف من تباين حجم التعبئة داخله. وتصنيف الشركة "
             "متماسكٌ سلوكياً لا اسماً: خط سفر ديزله ٤٨٪ وسلته ٣٦ لتراً، والحيوية "
             "ديزلها ١٨٪. ولذلك يُقارَن كلٌّ بشكله لا بمتوسط الشبكة.")
    return ws


# ═══════════════════════════════════════════════ ② المحطات — جدول البحث
def stations(wb):
    global TBL, S_ST_RANGE
    NC = 20
    ws = sheet(wb, S_ST, [3, 26, 8, 9, 9, 10, 11, 9, 10, 8, 8, 7, 10, 10, 12, 12,
                          13, 26, 8, 10],
               "المحطات — جدول واحد تقرأ منه الحاسبة وكل الأوراق",
               f"{TOT['n']} محطة مقيسة من كاش إن · {NET['days']:,} يوم-محطة · "
               "مرتَّبة بقيمة الفرصة · والبوابة في آخر عمود", NC, freeze="D5")
    r = 4
    W.header(ws, r, ["", "المحطة", "الرمز", "المنطقة", "التصنيف", "زيارة/يوم",
                     "لتر/يوم", "لتر/زيارة", "الفاتورة", "ديزل٪", "ليلي٪",
                     "ذروة", "فجوة الذروة", "فجوة السلة", "قيمة المعاملات",
                     "قيمة السلة", "الإجمالي ر/سنة", "البوابات والحالات",
                     "عمّال", "حصة المسائية"])
    ws.row_dimensions[r].height = 34
    r += 1
    first = r
    for s in ST:
        gate = any(c_ in s["conds"] for c_ in
                   ("فجوة مطابقة", "تغيير مسار", "بيانات دفع مفقودة"))
        fill = PatternFill("solid", fgColor=B.T_GOLD) if gate else None
        txt(ws, r, 2, s["name"], align=W.RGT, font=W.BOLD, fill=fill)
        txt(ws, r, 3, s["code"], font=Font(name=B.FONT, size=9, color=B.INK2), fill=fill)
        txt(ws, r, 4, s["region"], font=Font(name=B.FONT, size=9, color=B.INK2), fill=fill)
        txt(ws, r, 5, s["seg"], font=Font(name=B.FONT, size=9, color=B.INK2), fill=fill)
        txt(ws, r, 6, s["vpd"], fmt=NUM, fill=fill or W.CALC)
        txt(ws, r, 7, s["lpd"], fmt=NUM, fill=fill or W.CALC)
        txt(ws, r, 8, s["lpv"], fmt=NUM2, fill=fill or W.CALC)
        txt(ws, r, 9, s["inv"], fmt=NUM2, fill=fill or W.CALC)
        txt(ws, r, 10, s["diesel"], fmt=PCT, fill=fill or W.CALC)
        txt(ws, r, 11, s["night"], fmt=PCT, fill=fill or W.CALC)
        txt(ws, r, 12, s["peak"], fmt=NUM, fill=fill or W.CALC)
        txt(ws, r, 13, s["gap_peak"], fmt=NUM2, fill=fill or W.CALC)
        txt(ws, r, 14, s["gap_fill"], fmt=NUM2, fill=fill or W.CALC)
        txt(ws, r, 15, s["sar_txn"], fmt=MONEY, fill=fill or W.CALC)
        txt(ws, r, 16, s["sar_fill"], fmt=MONEY, fill=fill or W.CALC)
        txt(ws, r, 17, f"=O{r}+P{r}", fmt=MONEY, fill=fill or W.FS, font=W.BOLD)
        txt(ws, r, 18, ("⛔ " if gate else "") + " · ".join(s["conds"]), align=W.WRAP,
            fill=fill, font=Font(name=B.FONT, size=8,
                                 color=B.D_GOLD if gate else B.INK3))
        wk, evs = WORKERS.get(s["code"], (0, 0))
        txt(ws, r, 19, wk or "—", fmt=NUM if wk else None, fill=fill or W.CALC)
        txt(ws, r, 20, evs or "—", fmt=PCT if evs else None, fill=fill or W.CALC)
        r += 1
    last = r - 1
    txt(ws, r, 2, "الإجمالي", align=W.RGT, font=W.BOLD, fill=W.FS)
    for c_ in (3, 4, 5): txt(ws, r, c_, "", fill=W.FS)
    for c_ in (6, 7): txt(ws, r, c_, f"=SUM({get_column_letter(c_)}{first}:{get_column_letter(c_)}{last})",
                          fmt=NUM, fill=W.FS, font=W.BOLD)
    for c_ in (8, 9, 10, 11, 12): txt(ws, r, c_, "", fill=W.FS)
    for c_ in (13, 14): txt(ws, r, c_, f"=SUM({get_column_letter(c_)}{first}:{get_column_letter(c_)}{last})",
                            fmt=NUM, fill=W.FS, font=W.BOLD)
    for c_ in (15, 16, 17): txt(ws, r, c_, f"=SUM({get_column_letter(c_)}{first}:{get_column_letter(c_)}{last})",
                                fmt=MONEY, fill=W.FS, font=W.BOLD)
    txt(ws, r, 18, "", fill=W.FS)
    txt(ws, r, 19, f"=SUM(S{first}:S{last})", fmt=NUM, fill=W.FS, font=W.BOLD)
    txt(ws, r, 20, SG["eve_txn"], fmt=PCT, fill=W.FS, font=W.BOLD)
    r += 1
    for c_, col in (("M", 13), ("Q", 17), ("F", 6)):
        ws.conditional_formatting.add(f"{c_}{first}:{c_}{last}",
                                      DataBarRule(start_type="num", start_value=0,
                                                  end_type="max", color=B.ORANGE,
                                                  showValue=True))
    ws.auto_filter.ref = f"B{first-1}:T{last}"
    r = note(ws, r, NC,
             "⛔ الصفّ الذهبي = بوابة مفتوحة (فجوة مطابقة · تغيير مسار · بيانات دفع "
             "مفقودة). «فجوة الذروة» معاملة/يوم مقارنةً بشكل ساعات تصنيف المحطة، "
             "و«فجوة السلة» لتراً لكل زيارة دون ما يفسّره مزيجها. "
             "والقيمة تُحتسب بهامش مرجَّح بمزيج كل محطة لا بمتوسط الشبكة.",
             fill=PatternFill("solid", fgColor=B.T_GOLD),
             font=Font(name=B.FONT, size=9, bold=True, color=B.D_GOLD), h=36)
    TBL = f"'{S_ST}'!$B${first}:$T${last}"
    S_ST_RANGE = f"'{S_ST}'!$B${first}:$B${last}"
    return ws


# ═══════════════════════════════════════════════ ③ شكل الساعات — بيانات البحث
def hourshapes(wb):
    """حصة كل ساعة من معاملات اليوم — لكل محطة، ثم وسيط كل تصنيف.
       منها تُشتقّ فجوة الزيارات في الحاسبة أمام القارئ لا خلف ستار."""
    global HRS, HREF
    NC = 29
    ws = sheet(wb, S_HR, [3, 24, 8, 9] + [6.2] * 24 + [26],
               "شكل الساعات — حصة كل ساعة من معاملات اليوم",
               f"{len(ST)} محطة من ملف ساعات الشبكة · ثم وسيط كل تصنيف في آخر الورقة "
               "· وهذه الورقة مرجعٌ تقرأ منه الحاسبة، لا ورقة عرض", NC, freeze="E5")
    r = 4
    W.header(ws, r, ["", "المحطة", "الرمز", "التصنيف"]
             + [f"{h:02d}" for h in range(24)] + ["المجموع"])
    r += 1
    first = r
    prof = {}
    for s in ST:
        v = HOURS_RAW.get(s["code"])
        t = sum(v) if v else 0
        sh = [x / t for x in v] if t else [0.0] * 24
        prof.setdefault(s["seg"], []).append(sh)
        txt(ws, r, 2, s["name"], align=W.RGT, font=W.BOLD)
        txt(ws, r, 3, s["code"], font=Font(name=B.FONT, size=8, color=B.INK2))
        txt(ws, r, 4, s["seg"], font=Font(name=B.FONT, size=8, color=B.INK2))
        for h in range(24):
            txt(ws, r, 5 + h, sh[h], fmt=PCT2, fill=W.CALC,
                font=Font(name=B.FONT, size=8, color=B.INK))
        txt(ws, r, 29, f"=SUM(E{r}:AB{r})", fmt=PCT, fill=W.FS, font=W.BOLD)
        r += 1
    last = r - 1
    ws.conditional_formatting.add(f"E{first}:AB{last}",
                                  DataBarRule(start_type="num", start_value=0,
                                              end_type="max", color=B.GOLD,
                                              showValue=True))
    r += 1
    r = W.band(ws, r, NC, "وسيط كل تصنيف — وهو الشكل الذي تُقاس عليه الفجوة")
    W.header(ws, r, ["", "التصنيف", "محطات", ""]
             + [f"{h:02d}" for h in range(24)] + ["المجموع"])
    r += 1
    rfirst = r
    import statistics as stx
    for seg in [s["seg"] for s in SEGS]:
        rows = prof.get(seg, [])
        med = [stx.median([p[h] for p in rows]) for h in range(24)] if rows else [0] * 24
        HREF[seg] = med
        txt(ws, r, 2, seg, align=W.RGT, font=W.BOLD,
            fill=PatternFill("solid", fgColor=B.T_ORANGE))
        txt(ws, r, 3, len(rows), fmt=NUM, fill=PatternFill("solid", fgColor=B.T_ORANGE))
        txt(ws, r, 4, "وسيط", fill=PatternFill("solid", fgColor=B.T_ORANGE),
            font=Font(name=B.FONT, size=8, color=B.INK2))
        for h in range(24):
            txt(ws, r, 5 + h, med[h], fmt=PCT2,
                fill=PatternFill("solid", fgColor=B.T_ORANGE),
                font=Font(name=B.FONT, size=8, color=B.INK))
        txt(ws, r, 29, f"=SUM(E{r}:AB{r})", fmt=PCT, fill=W.FS, font=W.BOLD)
        r += 1
    rlast = r - 1
    r = note(ws, r, NC,
             "الوسيط لا يجمع ١٠٠٪ بالضبط — لأنه وسيطُ كل ساعةٍ على حدة لا توزيعُ "
             "محطةٍ واحدة. وهذا مقصود: نقارن كل ساعة بنظيرتها، ولا نفترض أن محطةً "
             "بعينها هي المعيار. ونصف محطات كل تصنيف فوق وسيطه بالتعريف.",
             fill=PatternFill("solid", fgColor=B.T_GOLD),
             font=Font(name=B.FONT, size=9, bold=True, color=B.D_GOLD))
    HRS = dict(first=first, last=last, rfirst=rfirst, rlast=rlast)
    return ws


# ═══════════════════════════════════════════════ ④ حاسبة المبيعات
def calculator(wb):
    NC = 12
    ws = sheet(wb, S_CALC,
               [3, 30, 14, 14, 14, 4, 13, 13, 4, 15, 15, 40],
               "حاسبة المبيعات — سلسلةٌ واحدة من خمس خطوات",
               "اختر المحطة، ثم اقرأ من أعلى إلى أسفل: قيمة اللتر ← الزيادة في "
               "الزيارات ← مستهدف العامل ← القيمة والكلفة ← بطاقة التسليم · "
               "الذهبيّ مُدخَلات وما عداه صيغٌ حيّة", NC)
    H_ = f"'{S_HR}'"
    SREF = f"{H_}!$E${HRS['rfirst']}:$AB${HRS['rlast']}"
    SNAM = f"{H_}!$B${HRS['rfirst']}:$B${HRS['rlast']}"
    SSTA = f"{H_}!$E${HRS['first']}:$AB${HRS['last']}"
    SCOD = f"{H_}!$B${HRS['first']}:$B${HRS['last']}"
    r = 4
    r = note(ws, r, NC,
             "كل رقم في هذه الورقة مشتقٌّ أمامك: لا رقم مُدخَل إلا ما كان بخلفية "
             "ذهبية، وكل ما عداه يقول من أين جاء في عمود «من أين». "
             "وإن غيّرت المحطة تغيّرت السلسلة كلها.",
             fill=PatternFill("solid", fgColor=B.T_ORANGE),
             font=Font(name=B.FONT, size=10, bold=True, color=B.ORANGE))
    r += 1

    # ── ⓪ المحطة
    r = W.band(ws, r, NC, "⓪ المحطة")
    txt(ws, r, 2, "المحطة", align=W.RGT,
        font=Font(name=B.FONT, size=12, bold=True, color=B.ORANGE))
    merge(ws, r, 3, 5, DEFAULT_ST, fill=W.FI,
          font=Font(name=B.FONT, size=12, bold=True, color=B.BLUE))
    txt(ws, r, 6, "", box=False)
    merge(ws, r, 7, 12, "◀ اضغط الخلية ثم اختر من السهم — "
          f"{TOT['n']} محطة مقيسة من كاش إن", align=W.RGT,
          font=Font(name=B.FONT, size=9, color=B.INK3))
    ws.row_dimensions[r].height = 24
    PICK = f"$C${r}"
    dv = DataValidation(type="list", formula1=f"{S_ST_RANGE}", allow_blank=False)
    ws.add_data_validation(dv); dv.add(ws.cell(r, 3))
    r += 1
    W.header(ws, r, ["", "من بيانات المحطة", "القيمة", "", "",
                     "", "من بيانات المحطة", "القيمة", "", "", "", "المصدر"])
    r += 1
    A = {}
    pairs = [("الرمز", 2, None, "كاش إن"), ("المنطقة", 3, None, "كاش إن"),
             ("التصنيف", 4, None, "ملف التحليلات — تصنيف الشركة"),
             ("زيارة/يوم", 5, NUM, "كاش إن — معدّل النافذة المقيسة"),
             ("لتر لكل زيارة", 7, NUM2, "كاش إن"),
             ("حصة الديزل من اللترات", 9, PCT, "كاش إن"),
             ("عدد العمّال", None, NUM, "تقرير العمّال والورديات"),
             ("حصة الوردية المسائية من المعاملات", None, PCT,
              "تقرير العمّال والورديات")]
    half = 4
    for i in range(half):
        for side in (0, 1):
            idx = i + side * half
            lab, col, fm, src = pairs[idx]
            bc, vc, sc = (2, 3, 5) if side == 0 else (7, 8, 12)
            txt(ws, r + i, bc, lab, align=W.RGT, font=W.BOLD)
            if col is None:                 # العمّال وحصة المسائية في آخر عمودين
                col = 18 if lab == "عدد العمّال" else 19
            f = f'=VLOOKUP({PICK},{TBL},{col},FALSE)'
            txt(ws, r + i, vc, f, fmt=fm, fill=W.CALC,
                font=Font(name=B.FONT, size=10, color=B.GOOD))
            A[lab] = f"${get_column_letter(vc)}${r+i}"
            if side == 0:
                merge(ws, r + i, 4, 5, src, align=W.RGT,
                      font=Font(name=B.FONT, size=8, color=B.INK3))
                txt(ws, r + i, 6, "", box=False)
            else:
                merge(ws, r + i, 9, 12, src, align=W.RGT,
                      font=Font(name=B.FONT, size=8, color=B.INK3))
    r += half
    VPD, LPV, DSL = A["زيارة/يوم"], A["لتر لكل زيارة"], A["حصة الديزل من اللترات"]
    SEG, WK = A["التصنيف"], A["عدد العمّال"]
    EVS = A["حصة الوردية المسائية من المعاملات"]
    r += 1

    # ── ① قيمة اللتر
    r = W.band(ws, r, NC, "① قيمة اللتر في هذه المحطة — بالضبط كما حسبناها")
    W.header(ws, r, ["", "الخطوة", "القيمة", "الوحدة", "", "", "", "", "",
                     "", "", "من أين جاءت"])
    r += 1
    L0 = r
    steps = [
        ("سعر المضخة — بنزين ٩١", PRICE_["بنزين ٩١"], MONEY2, "ريال/لتر", "inp",
         "سعرٌ فعلي من مبيعاتنا · شامل الضريبة"),
        ("سعر المضخة — ديزل", PRICE_["ديزل"], MONEY2, "ريال/لتر", "inp",
         "سعرٌ فعلي من مبيعاتنا · شامل الضريبة"),
        ("هامش لتر البنزين", K["margin_petrol"], RIYAL4, "ريال/لتر", "inp",
         "⛔ انحدار هامش ١٩ محطة على حصة ديزلها · R²=٠٫٥٦ — لا جدول معتمد"),
        ("هامش لتر الديزل", K["margin_diesel"], RIYAL4, "ريال/لتر", "inp",
         "⛔ الطرف الآخر للانحدار · أقصى حصة ديزل مرصودة ٤٣٪ فهو استقراء"),
        ("حصة الديزل في هذه المحطة", f"={DSL}", PCT, "٪ من اللترات", "calc",
         "من لترات المحطة نفسها — لا من متوسط الشبكة"),
        ("الهامش المرجَّح لهذه المحطة", None, RIYAL4, "ريال/لتر", "key",
         "الصيغة: هامش البنزين × (١−حصة الديزل) + هامش الديزل × حصة الديزل"),
        ("وبالهللة", None, HAL, "هللة/لتر", "key", "الرقم نفسه بالهللة"),
        ("كلفة التشغيل بالأساس المختار", None, HAL, "هللة/لتر", "calc",
         "من جدول الأسُس أدناه — والافتراضي «تشغيل المحطة»"),
        ("صافي اللتر", None, HAL, "هللة/لتر", "key",
         "الصيغة: الهامش المرجَّح − كلفة التشغيل · وهو ما يبقى من كل لتر إضافي"),
    ]
    for i, (lab, v, fm, u, kind, src) in enumerate(steps):
        rr = r + i
        bold = kind == "key"
        fill = (W.FI if kind == "inp" else
                PatternFill("solid", fgColor=B.T_GOOD) if bold else W.CALC)
        txt(ws, rr, 2, lab, align=W.RGT,
            font=Font(name=B.FONT, size=11 if bold else 10, bold=True,
                      color=B.INK))
        if kind == "inp":
            W.inp(ws, rr, 3, fm, v)
        else:
            f = v if isinstance(v, str) else {
                5: f"=$C${L0+2}*(1-$C${L0+4})+$C${L0+3}*$C${L0+4}",
                6: f"=$C${L0+5}*100",
                7: 0,                       # تُربط بجدول الأسُس بعد تعريفه
                8: f"=$C${L0+6}-$C${L0+7}",
            }[i]
            txt(ws, rr, 3, f, fmt=fm, fill=fill,
                font=Font(name=B.FONT, size=12 if bold else 10, bold=bold,
                          color=B.D_GOOD if bold else B.INK))
        txt(ws, rr, 4, u, font=Font(name=B.FONT, size=9, color=B.INK3))
        txt(ws, rr, 5, ""); txt(ws, rr, 6, "", box=False)
        gate = src.startswith("⛔")
        merge(ws, rr, 7, 12, src, align=W.RGT,
              fill=PatternFill("solid", fgColor=B.T_GOLD) if gate else None,
              font=Font(name=B.FONT, size=8, bold=gate,
                        color=B.D_GOLD if gate else B.INK3))
    BLEND = f"$C${L0+5}"
    NETL = f"$C${L0+8}"
    PRICEW = f"($C${L0}*(1-{DSL})+$C${L0+1}*{DSL})"
    r += len(steps)
    r = note(ws, r, NC,
             "وهذا هو «قيمة اللتر» بعينها: هامش المساهمة — الإيراد ناقص شراء الوقود، "
             "قبل أي مصروف تشغيل. وهو وحده ما يتحرّك باللتر الإضافي. "
             "وسعر المضخة معروضٌ للسياق لا للحساب: الهامش لا يُشتقّ منه بل من "
             "قائمة الدخل.")
    r += 1

    r = W.band(ws, r, NC, "أسُس الكلفة الثلاثة — اختر الأساس الذي يخصّ قرارك")
    W.header(ws, r, ["", "الأساس", "هللة/لتر", "", "", "", "", "",
                     "", "", "", "يصلح لقرار"])
    r += 1
    LUT = r
    for b in C["bases"]:
        txt(ws, r, 2, b["base"], align=W.RGT)
        txt(ws, r, 3, b["cpl"] * 100, fmt=HAL, fill=W.CALC)
        for c_ in (4, 5): txt(ws, r, c_, "")
        txt(ws, r, 6, "", box=False)
        merge(ws, r, 7, 12, b["use"], align=W.RGT,
              font=Font(name=B.FONT, size=9, color=B.INK2))
        r += 1
    LUTE = r - 1
    txt(ws, r, 2, "الأساس المختار", align=W.RGT,
        font=Font(name=B.FONT, size=11, bold=True, color=B.ORANGE))
    merge(ws, r, 3, 5, "تشغيل المحطة — محمَّل بالكامل", fill=W.FI,
          font=Font(name=B.FONT, size=10, bold=True, color=B.BLUE))
    BASIS = f"$C${r}"
    txt(ws, r, 6, "", box=False)
    merge(ws, r, 7, 12, "◀ غيّره ليتغيّر «صافي اللتر» أعلاه وكل ما بُني عليه",
          align=W.RGT, font=Font(name=B.FONT, size=9, color=B.INK3))
    dv1 = DataValidation(type="list",
                         formula1=f'"{",".join(b["base"] for b in C["bases"])}"',
                         allow_blank=False)
    ws.add_data_validation(dv1); dv1.add(ws.cell(r, 3))
    r += 1
    # اربط صيغة كلفة التشغيل بجدول الأسُس بعد أن عُرفت أسطره
    ws.cell(L0 + 7, 3).value = (f"=VLOOKUP({BASIS},$B${LUT}:$C${LUTE},2,FALSE)")
    r = note(ws, r, NC,
             "الأساس الحدّي صفرٌ لأن النقص في ساعات الذروة طاقةُ خدمةٍ لا طاقة ضخّ: "
             "المضخّات والمرافق والأجور مدفوعة أصلاً، فاللتر الإضافي على الطاقم "
             "القائم لا يضيف كلفةً مقيسة. ومن يقرّر فتح وردية أو توظيفاً يستعمل "
             "الأساس الثاني، ومن يقرّر رأس مال جديداً يستعمل الثالث.")
    r += 1

    # ── ② الزيادة في الزيارات
    r = W.band(ws, r, NC,
               "② الزيادة في الزيارات لهذه المحطة — ساعةً ساعة، لا رقماً مُنزَلاً")
    r = note(ws, r, NC,
             "القاعدة: نقارن حصة كل ساعة عندنا بحصتها عند وسيط تصنيفنا. وما نحن "
             "دونه في ساعات ذروتنا الثماني وحدها هو الفجوة المؤكَّدة — أمّا النقص "
             "في ساعات الهدوء فاختلاف طلبٍ لا فاقد، فلا يُطالَب به أحد.",
             fill=PatternFill("solid", fgColor=B.T_NEUTRAL))
    W.header(ws, r, ["", "الساعة", "حصتنا", "فعلي/يوم", "شكل التصنيف",
                     "المتوقَّع/يوم", "", "الفجوة", "", "ذروة؟", "الوردية",
                     "القراءة"])
    r += 1
    H0 = r
    for h in range(24):
        rr = r + h
        txt(ws, rr, 2, f"{h:02d}:٠٠", font=W.BOLD)
        txt(ws, rr, 3,
            f"=INDEX({SSTA},MATCH({PICK},{SCOD},0),{h+1})", fmt=PCT2, fill=W.CALC)
        txt(ws, rr, 4, f"=C{rr}*{VPD}", fmt=NUM, fill=W.CALC)
        txt(ws, rr, 5,
            f"=INDEX({SREF},MATCH({SEG},{SNAM},0),{h+1})", fmt=PCT2, fill=W.CALC)
        txt(ws, rr, 6, f"=E{rr}*{VPD}", fmt=NUM, fill=W.CALC)
        txt(ws, rr, 7, "", box=False)
        txt(ws, rr, 8, f'=IF(J{rr}="★",MAX(0,F{rr}-D{rr}),0)', fmt=NUM2,
            fill=W.CALC, font=W.BOLD)
        txt(ws, rr, 9, "", box=False)
        txt(ws, rr, 10, f'=IF(C{rr}>=LARGE($C${H0}:$C${H0+23},8),"★","")',
            font=Font(name=B.FONT, size=10, color=B.ORANGE))
        txt(ws, rr, 11, "صباحية" if h < 12 else "مسائية",
            font=Font(name=B.FONT, size=9, color=B.INK2))
        txt(ws, rr, 12, "", align=W.WRAP,
            font=Font(name=B.FONT, size=8, color=B.D_BAD))
    H1 = H0 + 23
    for rr in range(H0, H1 + 1):
        ws.cell(rr, 12).value = (f'=IF(H{rr}>0,"فجوة تُحتسب — ساعة ذروة ونحن '
                                 f'دون الشكل","")')
    ws.conditional_formatting.add(f"H{H0}:H{H1}", CellIsRule(
        operator="greaterThan", formula=["0"],
        fill=PatternFill("solid", fgColor=B.T_BAD),
        font=Font(name=B.FONT, size=10, bold=True, color=B.D_BAD)))
    ws.conditional_formatting.add(f"D{H0}:D{H1}",
                                  DataBarRule(start_type="num", start_value=0,
                                              end_type="max", color=B.GOLD,
                                              showValue=True))
    r = H1 + 1
    rows2 = [
        ("فجوة ساعات الذروة — الرقم المؤكَّد", f"=SUM(H{H0}:H{H1})", NUM2,
         "معاملة/يوم", "key",
         "مجموع الفجوة في ساعات ذروتنا الثماني وحدها"),
        ("منها في الوردية الصباحية ٠٠–١١", f"=SUM(H{H0}:H{H0+11})", NUM2,
         "معاملة/يوم", "calc", "الساعات ٠٠ إلى ١١"),
        ("ومنها في الوردية المسائية ١٢–٢٣", f"=SUM(H{H0+12}:H{H1})", NUM2,
         "معاملة/يوم", "calc", "الساعات ١٢ إلى ٢٣"),
        ("نسبة الإغلاق خلال ستة أشهر", K["close"], PCT, "٪", "inp",
         "سياسة الخطة — ٤٠٪ في كل محطات الشبكة"),
        ("المطلوب في اليوم", None, NUM2, "معاملة/يوم", "key",
         "الصيغة: فجوة الذروة × نسبة الإغلاق"),
    ]
    G0 = r
    for i, (lab, v, fm, u, kind, src) in enumerate(rows2):
        rr = r + i
        bold = kind == "key"
        fill = (W.FI if kind == "inp" else
                PatternFill("solid", fgColor=B.T_GOOD) if bold else W.FS)
        txt(ws, rr, 2, lab, align=W.RGT,
            font=Font(name=B.FONT, size=11 if bold else 10, bold=True))
        if kind == "inp":
            W.inp(ws, rr, 3, fm, v)
        else:
            f = v if v is not None else f"=$C${G0}*$C${G0+3}"
            txt(ws, rr, 3, f, fmt=fm, fill=fill,
                font=Font(name=B.FONT, size=12 if bold else 10, bold=bold,
                          color=B.D_GOOD if bold else B.INK))
        txt(ws, rr, 4, u, font=Font(name=B.FONT, size=9, color=B.INK3))
        txt(ws, rr, 5, ""); txt(ws, rr, 6, "", box=False)
        merge(ws, rr, 7, 12, src, align=W.RGT,
              font=Font(name=B.FONT, size=8, color=B.INK3))
    GAPP = f"$C${G0}"; GAPM = f"$C${G0+1}"; GAPE = f"$C${G0+2}"
    CLOSE = f"$C${G0+3}"; NEED = f"$C${G0+4}"
    r += len(rows2)
    r = note(ws, r, NC,
             "«ذروة؟» تُحدَّد بشكل المحطة نفسها: ساعاتها الثماني الأعلى حصةً. "
             "ولهذا قد تسقط ساعةٌ فجوتُها كبيرة لأنها ليست من ساعاتنا الأعلى — "
             "وهو تحفّظٌ مقصود يجعل الرقم أصغر مما قد يكون، لا أكبر. "
             "وإن تساوت ساعتان عند الحدّ الثامن دخلتا معاً.")
    r += 1

    # ── ③ مستهدف الوردية والعامل
    r = W.band(ws, r, NC, "③ مستهدف الوردية والعامل — الرقم الذي يُسلَّم فعلاً")
    W.header(ws, r, ["", "الخطوة", "القيمة", "الوحدة", "", "", "", "", "",
                     "", "", "من أين جاءت"])
    r += 1
    T0 = r
    trows = [
        ("عدد عمّال المحطة", f"={WK}", NUM, "عامل", "calc",
         "تقرير العمّال والورديات — إجمالي المحطة"),
        ("حصة الوردية المسائية من المعاملات", f"={EVS}", PCT, "٪", "calc",
         "معاملات المساء ÷ معاملات اليوم — مقيسة"),
        ("⛔ عدد عمّال الوردية المسائية", None, NUM, "عامل", "inp",
         "الافتراضي هو التوزيع المتوازن مع الطلب — اكتب العدد المعتمد من "
         "العمليات فوقه، فهذا هو المُدخَل الملزم الوحيد هنا"),
        ("المطلوب من الوردية المسائية", None, NUM2, "معاملة/يوم", "calc",
         "الصيغة: المطلوب في اليوم × حصة المسائية من الفجوة"),
        ("المستهدف لكل عامل مسائي", None, NUM2, "معاملة/يوم", "key",
         "الصيغة: المطلوب من الوردية ÷ عدد عمّالها"),
        ("وبالساعة", None, NUM2, "معاملة/ساعة", "key",
         "على وردية اثنتي عشرة ساعة"),
        ("حِمله اليوم", None, NUM, "معاملة/يوم", "calc",
         "الصيغة: معاملات المساء ÷ عمّال المساء — للسياق"),
        ("الزيادة المطلوبة على حِمله", None, PCT, "٪", "key",
         "وهذا هو حجم الطلب الحقيقي من العامل"),
    ]
    for i, (lab, v, fm, u, kind, src) in enumerate(trows):
        rr = r + i
        bold = kind == "key"
        gate = lab.startswith("⛔")
        fill = (W.FI if kind == "inp" else
                PatternFill("solid", fgColor=B.T_GOOD) if bold else W.CALC)
        txt(ws, rr, 2, lab, align=W.RGT,
            font=Font(name=B.FONT, size=11 if bold else 10, bold=True,
                      color=B.D_GOLD if gate else B.INK))
        if kind == "inp":
            W.inp(ws, rr, 3, fm, None)
            ws.cell(rr, 3).value = f"=ROUND($C${T0}*$C${T0+1},0)"
        else:
            f = v if v is not None else {
                3: f"={NEED}*IF({GAPP}=0,0,{GAPE}/{GAPP})",
                4: f"=IF($C${T0+2}=0,0,$C${T0+3}/$C${T0+2})",
                5: f"=$C${T0+4}/12",
                6: f"=IF($C${T0+2}=0,0,{VPD}*$C${T0+1}/$C${T0+2})",
                7: f"=IF($C${T0+6}=0,0,$C${T0+4}/$C${T0+6})",
            }[i]
            txt(ws, rr, 3, f, fmt=fm, fill=fill,
                font=Font(name=B.FONT, size=12 if bold else 10, bold=bold,
                          color=B.D_GOOD if bold else B.INK))
        txt(ws, rr, 4, u, font=Font(name=B.FONT, size=9, color=B.INK3))
        txt(ws, rr, 5, ""); txt(ws, rr, 6, "", box=False)
        merge(ws, rr, 7, 12, src, align=W.RGT,
              fill=PatternFill("solid", fgColor=B.T_GOLD) if gate else None,
              font=Font(name=B.FONT, size=8, bold=gate,
                        color=B.D_GOLD if gate else B.INK3))
        if gate: ws.row_dimensions[rr].height = 26
    PERW = f"$C${T0+4}"; PERH = f"$C${T0+5}"; LOADE = f"$C${T0+6}"
    UPW = f"$C${T0+7}"; EVW = f"$C${T0+2}"
    r += len(trows)
    r = note(ws, r, NC,
             "⛔ بوابة: عدد عمّال الوردية هو المُدخَل الوحيد الذي لا نملكه مقيساً "
             "لكل محطة — وثلاثة ملفات تعطي للعمرة وحدها ٢٠ و٢٦ و٣٢. "
             "فالافتراضي هنا توزيعٌ متوازن مع الطلب، ولا يُحمَّل مستهدفٌ على وردية "
             "قبل أن تكتب العمليات العدد المعتمد في الخلية الذهبية.",
             fill=PatternFill("solid", fgColor=B.T_GOLD),
             font=Font(name=B.FONT, size=9, bold=True, color=B.D_GOLD), h=36)
    r += 1

    # ── ④ القيمة والكلفة
    r = W.band(ws, r, NC, "④ القيمة والكلفة — الخطوتان ① و② مضروبتان ببعضهما")
    W.header(ws, r, ["", "السطر", "في اليوم", "في السنة", "الوحدة", "",
                     "", "", "", "", "", "الصيغة بالكلمات"])
    r += 1
    V0 = r
    vrows = [
        ("لترات إضافية", f"={NEED}*{LPV}", "لتر", NUM2, NUM, "calc",
         "الصيغة: المطلوب في اليوم × لتر لكل زيارة"),
        ("مبيعات إضافية", f"=C{V0}*{PRICEW}", "ريال", MONEY2, MONEY, "calc",
         "الصيغة: اللترات × سعر البيع المرجَّح"),
        ("هامش مساهمة", f"=C{V0}*{BLEND}", "ريال", MONEY2, MONEY, "good",
         "الصيغة: اللترات × الهامش المرجَّح من الخطوة ①"),
        ("ناقص كلفة التشغيل", f"=-C{V0}*$C${L0+7}/100", "ريال", MONEY2, MONEY, "bad",
         "الصيغة: اللترات × كلفة التشغيل بالأساس المختار"),
        ("صافي المستهدف", f"=C{V0+2}+C{V0+3}", "ريال", MONEY2, MONEY, "key",
         "الصيغة: اللترات × صافي اللتر"),
        ("قيمة إغلاق الفجوة كاملةً", f"={GAPP}*{LPV}*{NETL}/100", "ريال",
         MONEY2, MONEY, "calc", "بلا نسبة إغلاق — الفجوة كلها"),
        ("قيمة معاملةٍ واحدة إضافية في اليوم", f"={LPV}*{NETL}/100", "ريال",
         MONEY2, MONEY, "good", "الصيغة: لتر لكل زيارة × صافي اللتر"),
    ]
    tf = {"good": B.T_GOOD, "bad": B.T_BAD, "key": B.T_BAND}
    tc = {"good": B.D_GOOD, "bad": B.D_BAD, "key": B.INK}
    for i, (lab, f1, u, fm1, fm2, tone, why) in enumerate(vrows):
        rr = r + i
        bold = tone == "key"
        fill = PatternFill("solid", fgColor=tf.get(tone, B.T_NEUTRAL))
        txt(ws, rr, 2, lab, align=W.RGT,
            font=Font(name=B.FONT, size=11 if bold else 10, bold=True))
        txt(ws, rr, 3, f1, fmt=fm1, fill=fill,
            font=Font(name=B.FONT, size=12 if bold else 10, bold=bold,
                      color=tc.get(tone, B.INK)))
        txt(ws, rr, 4, f"=C{rr}*365", fmt=fm2, fill=fill,
            font=Font(name=B.FONT, size=12 if bold else 10, bold=bold,
                      color=tc.get(tone, B.INK)))
        txt(ws, rr, 5, u, font=Font(name=B.FONT, size=9, color=B.INK3))
        txt(ws, rr, 6, "", box=False)
        merge(ws, rr, 7, 12, why, align=W.RGT,
              font=Font(name=B.FONT, size=8, color=B.INK3))
    NETD = f"$C${V0+4}"
    r += len(vrows)
    r = note(ws, r, NC,
             "ولاحظ أن «صافي المستهدف» يتبع الأساس الذي اخترته في الخطوة ①: "
             "بالأساس الحدّي هو الهامش كاملاً، وبالمحمَّل الكامل قد يخرج سالباً "
             "في محطاتٍ هامشها دون كلفتها محمَّلةً. وذلك ليس سبباً لرفض اللتر "
             "الإضافي، بل سببٌ لئلّا يُبنى عليه قرار رأس مال.")
    r += 1

    # ── ⑤ بطاقة التسليم
    r = W.band(ws, r, NC, "⑤ بطاقة التسليم — ما يُكتب في ورقة الوردية")
    W.header(ws, r, ["", "البند", "القيمة", "", "", "", "البند", "القيمة",
                     "", "", "", "ملاحظة"])
    r += 1
    K0 = r
    card = [
        ("المحطة", f"={PICK}", None, "الوردية المستهدَفة", '="المسائية ١٢:٠٠–٢٣:٥٩"',
         None),
        ("عدد العمّال في الوردية", f"={EVW}", NUM,
         "المستهدف لكل عامل في اليوم", f"={PERW}", NUM2),
        ("وبالساعة", f"={PERH}", NUM2,
         "الزيادة على حِمله الحالي", f"={UPW}", PCT),
        ("الرقم الذي يُقاس",
         '="معاملات ساعات الذروة المسائية وحدها"', None,
         "خط الأساس", '="الأسابيع الأربعة السابقة في المحطة نفسها"', None),
        ("متى يُحكَم", '="بعد ستة أسابيع ثم كل شهر"', None,
         "حدّ الأثر المقبول",
         f'="أكثر من {D["campaigns"]["noise"]*100:.1f}٪ — تشتّت خط الأساس"', None),
        ("مالك التنفيذ", '="مدير المحطة"', None,
         "مالك القياس", '="التجاري · والعمليات للطاقم"', None),
    ]
    for i, (l1, f1, m1, l2, f2, m2) in enumerate(card):
        rr = r + i
        txt(ws, rr, 2, l1, align=W.RGT, font=W.BOLD)
        merge(ws, rr, 3, 5, f1, fmt=m1,
              fill=PatternFill("solid", fgColor=B.T_ORANGE),
              font=Font(name=B.FONT, size=11, bold=True, color=B.ORANGE))
        txt(ws, rr, 6, "", box=False)
        txt(ws, rr, 7, l2, align=W.RGT, font=W.BOLD)
        merge(ws, rr, 8, 11, f2, fmt=m2,
              fill=PatternFill("solid", fgColor=B.T_ORANGE),
              font=Font(name=B.FONT, size=11, bold=True, color=B.ORANGE))
        txt(ws, rr, 12, "", box=False)
        ws.row_dimensions[rr].height = 22
    r += len(card)
    r = note(ws, r, NC,
             "وما لا يُطالَب به في هذه البطاقة عمداً: لترات الفاتورة الواحدة. "
             "ثلثا تباينها مزيج وقود لا سلوك، فلا يملكها العامل ولا يُحاسَب عليها. "
             "والمستهدف معاملاتٌ يخدمها، وهو ما يملكه فعلاً.",
             fill=PatternFill("solid", fgColor=B.T_ORANGE),
             font=Font(name=B.FONT, size=9, bold=True, color=B.ORANGE))
    return ws

# ═══════════════════════════════════════════════ ④ الربع الأول والثاني
def _q2rows():
    """تقرير ربحية المحطات المشغّلة — ر١ مقابل ر٢ لكل محطة، والأجور والمرافق معها"""
    cost = {r["code"]: r for r in csv.DictReader(open(Q2COST, encoding="utf-8"))}
    ct = {c["code"]: c for c in CONTR["stations"]}

    def f(v):
        try: return float(v)
        except (TypeError, ValueError): return 0.0

    out = []
    for r in csv.DictReader(open("outlets/data/q2-stations.csv", encoding="utf-8")):
        if not f(r["q2_vol"]): continue
        c = cost.get(r["code"], {})
        k = ct.get(r["code"], {})
        out.append(dict(
            code=r["code"], name=r["name"], terms=r["terms"],
            q1_sales=f(r["q1_sales"]), q2_sales=f(r["q2_sales"]),
            q1_vol=f(r["q1_vol"]), q2_vol=f(r["q2_vol"]),
            q1_net=f(r["q1_net"]), q2_net=f(r["q2_net"]),
            q1_margin=f(r["q1_margin"]), q2_margin=f(r["q2_margin"]),
            q1_staff=f(r["q1_staff"]), q2_staff=f(r["q2_staff"]),
            q1_darb=f(r["q1_darb"]), q2_darb=f(r["q2_darb"]),
            q2_owner=f(r["q2_owner"]),
            wages=f(c.get("q2_wages")), util=f(c.get("q2_util")),
            # «حصّتنا» = act من نموذج العقود: في عقود «نسبة المالك» تُسجَّل حصّتنا
            #   صافيَ المحطة لا بنداً باسم درب، فلو أُخذ العمود الخام لظهرت صفراً
            act=k.get("act", f(r["q2_darb"])), std=k.get("std", 0.0),
            gap=k.get("gap", 0.0), kind=k.get("kind", "—")))
    return sorted(out, key=lambda x: x["gap"])


def quarters_net(wb):
    NC = 18
    rows = _q2rows()
    ws = sheet(wb, "ر١ و ر٢ للمحطات",
               [3, 20, 7, 12, 12, 8, 11, 11, 8, 10, 10, 8, 10, 9, 11, 11, 11, 26],
               "مبيعات الربع الأول والثاني — المحطات المشغّلة",
               f"{len(rows)} محطة من تقرير ربحية الربع الثاني · "
               "المبيعات والحجم والصافي والأجور · وحصّتنا مقارنةً بالعمولة المعيارية "
               "(٣ هللات بنزين · ١٫٥ ديزل)", NC, freeze="D5")
    r = 4
    W.header(ws, r, ["", "المحطة", "الرمز", "مبيعات ر١", "مبيعات ر٢", "التغيّر",
                     "حجم ر١ لتر", "حجم ر٢ لتر", "التغيّر", "صافي ر١", "صافي ر٢",
                     "هامش ر٢", "أجور ر٢", "مرافق ر٢", "حصّتنا ر٢", "المعياري ر٢",
                     "الفجوة ر٢", "شروط العقد ونوعها"])
    ws.row_dimensions[r].height = 36
    r += 1
    f0 = r
    for s in rows:
        bad = s["gap"] < -50000
        fill = PatternFill("solid", fgColor=B.T_BAD) if bad else None
        txt(ws, r, 2, s["name"], align=W.RGT, font=W.BOLD, fill=fill)
        txt(ws, r, 3, s["code"], font=Font(name=B.FONT, size=8, color=B.INK2), fill=fill)
        txt(ws, r, 4, s["q1_sales"], fmt=MONEY, fill=fill or W.CALC)
        txt(ws, r, 5, s["q2_sales"], fmt=MONEY, fill=fill or W.CALC)
        txt(ws, r, 6, f"=IF(D{r}=0,\"—\",E{r}/D{r}-1)", fmt=PCT, fill=fill or W.CALC)
        txt(ws, r, 7, s["q1_vol"], fmt=NUM, fill=fill or W.CALC)
        txt(ws, r, 8, s["q2_vol"], fmt=NUM, fill=fill or W.CALC)
        txt(ws, r, 9, f"=IF(G{r}=0,\"—\",H{r}/G{r}-1)", fmt=PCT, fill=fill or W.CALC)
        txt(ws, r, 10, s["q1_net"], fmt=MONEY, fill=fill or W.CALC)
        txt(ws, r, 11, s["q2_net"], fmt=MONEY, fill=fill or W.CALC)
        dash = Font(name=B.FONT, size=10, color=B.INK3)
        txt(ws, r, 12, s["q2_margin"] / 100 if s["q2_margin"] else "—",
            fmt=PCT if s["q2_margin"] else None, fill=fill or W.CALC,
            font=None if s["q2_margin"] else dash)
        txt(ws, r, 13, s["wages"] or "—", fmt=MONEY if s["wages"] else None,
            fill=fill or W.CALC, font=None if s["wages"] else dash)
        txt(ws, r, 14, s["util"] or "—", fmt=MONEY if s["util"] else None,
            fill=fill or W.CALC, font=None if s["util"] else dash)
        txt(ws, r, 15, s["act"], fmt=MONEY, fill=fill or W.CALC)
        txt(ws, r, 16, s["std"], fmt=MONEY, fill=fill or W.CALC)
        txt(ws, r, 17, f"=O{r}-P{r}", fmt=MONEY, fill=fill or W.CALC, font=W.BOLD)
        txt(ws, r, 18, f"{s['kind']} — {s['terms']}", align=W.WRAP, fill=fill,
            font=Font(name=B.FONT, size=8, color=B.INK3))
        r += 1
    f1 = r - 1
    txt(ws, r, 2, "الإجمالي", align=W.RGT, font=W.BOLD, fill=W.FS)
    txt(ws, r, 3, "", fill=W.FS)
    for c_ in (4, 5, 7, 8, 10, 11, 13, 14, 15, 16, 17):
        txt(ws, r, c_, f"=SUM({get_column_letter(c_)}{f0}:{get_column_letter(c_)}{f1})",
            fmt=MONEY if c_ not in (7, 8) else NUM, fill=W.FS, font=W.BOLD)
    txt(ws, r, 6, f"=E{r}/D{r}-1", fmt=PCT, fill=W.FS, font=W.BOLD)
    txt(ws, r, 9, f"=H{r}/G{r}-1", fmt=PCT, fill=W.FS, font=W.BOLD)
    txt(ws, r, 12, f"=K{r}/E{r}", fmt=PCT, fill=W.FS, font=W.BOLD)
    txt(ws, r, 18, "", fill=W.FS)
    tot = r
    r += 1
    for c_ in ("F", "I", "Q"): sign_cf(ws, f"{c_}{f0}:{c_}{f1}")
    ws.conditional_formatting.add(f"L{f0}:L{f1}",
                                  DataBarRule(start_type="num", start_value=0,
                                              end_type="max", color=B.GOOD, showValue=True))
    ws.auto_filter.ref = f"B{f0-1}:R{f1}"
    r = note(ws, r, NC,
             f"«المعياري» هو ما كانت حصّتنا ستكونه بعمولة درب المعيارية على حجم الربع "
             f"الثاني، و«الفجوة» الفرقُ — كلاهما للربع لا للسنة. "
             f"ومجموع الفجوة {sum(x['gap'] for x in rows):,.0f} للربع، أي "
             f"{CONTR['gap_year']:,.0f} ريال في السنة — أكبر رقمٍ في الخطة كلها، "
             "ولا يحتاج لتراً إضافياً ولا حملة: يحتاج جلسة تفاوض على العقود. "
             "والصفوف الحمراء هي التي تتجاوز فجوتها ٥٠ ألف ريال في الربع.",
             fill=PatternFill("solid", fgColor=B.T_GOLD),
             font=Font(name=B.FONT, size=9, bold=True, color=B.D_GOLD), h=44)
    r += 1

    # ملخص الأرباع
    r = W.band(ws, r, NC, "الخلاصة — الربع الثاني مقابل الأول على مستوى هذه المحطات")
    W.header(ws, r, ["", "المقياس", "الربع الأول", "الربع الثاني", "التغيّر", "",
                     "", "", "", "", "", "", "", "", "", "", "", "القراءة"])
    r += 1
    s0 = r
    agg = [
        ("المبيعات ريال", sum(s["q1_sales"] for s in rows),
         sum(s["q2_sales"] for s in rows), MONEY,
         "وهي مبيعات المحطات المشغّلة وحدها لا الشبكة"),
        ("الحجم لتر", sum(s["q1_vol"] for s in rows),
         sum(s["q2_vol"] for s in rows), NUM,
         "الحجم يهبط أكثر من المبيعات — أي أن السعر المتوسط ارتفع بالمزيج"),
        ("صافي المحطات ريال", sum(s["q1_net"] for s in rows),
         sum(s["q2_net"] for s in rows), MONEY, ""),
        ("عدد العمّال", sum(s["q1_staff"] for s in rows),
         sum(s["q2_staff"] for s in rows), NUM,
         "قلّ مع الحجم — فالعمالة تُضبط فعلاً"),
    ]
    gapq = sum(s["gap"] for s in rows)
    stdq = sum(s["std"] for s in rows)
    actq = sum(s["act"] for s in rows)
    for lab, a, b, fm, rd in agg:
        txt(ws, r, 2, lab, align=W.RGT, font=W.BOLD)
        txt(ws, r, 3, a, fmt=fm, fill=W.CALC)
        txt(ws, r, 4, b, fmt=fm, fill=W.CALC)
        txt(ws, r, 5, f"=D{r}/C{r}-1", fmt=PCT, fill=W.CALC, font=W.BOLD)
        for c_ in range(6, 18): txt(ws, r, c_, "", box=False)
        txt(ws, r, 18, rd, align=W.WRAP, font=Font(name=B.FONT, size=9, color=B.INK2))
        r += 1
    sign_cf(ws, f"E{s0}:E{r-1}")
    # سطر الفجوة — أساسه أساس الجدول أعلاه نفسه فيتطابق المجموعان
    txt(ws, r, 2, "فجوة العمولة عن المعياري", align=W.RGT,
        font=Font(name=B.FONT, size=10, bold=True, color=B.D_BAD))
    txt(ws, r, 3, actq, fmt=MONEY, fill=PatternFill("solid", fgColor=B.T_BAD),
        font=Font(name=B.FONT, size=10, bold=True, color=B.D_BAD))
    txt(ws, r, 4, stdq, fmt=MONEY, fill=PatternFill("solid", fgColor=B.T_BAD),
        font=Font(name=B.FONT, size=10, bold=True, color=B.D_BAD))
    txt(ws, r, 5, f"=C{r}-D{r}", fmt=MONEY,
        fill=PatternFill("solid", fgColor=B.T_BAD),
        font=Font(name=B.FONT, size=11, bold=True, color=B.D_BAD))
    for c_ in range(6, 18): txt(ws, r, c_, "", box=False)
    txt(ws, r, 18, "حصّتنا الفعلية مقابل المعياري — والعمودان هنا ليسا ر١ و ر٢ "
        "بل الفعلي والمعياري للربع الثاني", align=W.WRAP,
        font=Font(name=B.FONT, size=9, color=B.D_BAD))
    ws.row_dimensions[r].height = 26
    r += 1
    r = note(ws, r, NC,
             f"وفجوة الربع {gapq:,.0f} ريالاً × أربعة أرباع = "
             f"{gapq*4:,.0f} ريال في السنة — وهو الرقم المعروض في الملخص.",
             fill=PatternFill("solid", fgColor=B.T_GOLD),
             font=Font(name=B.FONT, size=9, bold=True, color=B.D_GOLD))
    r = note(ws, r, NC,
             "والقراءة الواحدة من الجدولين: الربع الثاني أضعف حجماً في أغلب المحطات، "
             "لكن الفارق الأكبر في صافينا ليس الحجم — هو بنود العقد. "
             "فمن يعالج هذا الربع بحملةٍ يعالج العَرَض.")
    return ws


# ═══════════════════════════════════════════════ ⑤ العمرة — بيان يومي
def mk007(wb):
    NC = 11
    ws = sheet(wb, S_MK, [3, 24, 13, 13, 13, 4, 12, 12, 12, 4, 28],
               f"{D['code']} · {D['name']} — البيان اليومي الوحيد لدينا",
               f"{D['source']} · {D['days']} يوماً كاملاً ({D['first']} → {D['last']}) · "
               f"{D['rows']:,} عملية · يوم {D['partial']} ناقص فاستُبعد · "
               "وهي المحطة الوحيدة التي نملك لها بياناً بالساعة", NC)
    r = 4
    r = note(ws, r, NC,
             "الحكم أولاً: هذه المحطة ليست في انحدار. مقابل يناير هي "
             f"{BL['vs_jan_t']*100:+.1f}٪ معاملةً و{BL['vs_jan_l']*100:+.1f}٪ لتراً، "
             f"ووسيط نظائرها في مكة {BL['peers_med']*100:+.1f}٪. "
             "أمّا الـ−٢٢٪ المتداولة فهي قياسٌ على ذروة مارس، وذروة مارس موسم لا "
             "خطَّ أساس. وما بقي بعد الفرز فجوةٌ واحدة: ساعات الذروة المسائية.",
             fill=PatternFill("solid", fgColor=B.T_GOLD),
             font=Font(name=B.FONT, size=10, bold=True, color=B.D_GOLD), h=46)
    r += 1

    # خطوط الأساس
    r = W.band(ws, r, NC, "① أي خط أساس تقيس عليه — هو ما يقرّر أترى أزمةً أم استواءً")
    W.header(ws, r, ["", "خط الأساس", "لتر/يوم", "معاملة/يوم", "", "",
                     "الفرق لتراً", "الفرق معاملةً", "", "", "القراءة"])
    r += 1
    b0 = r
    for i, b in enumerate(BL["rows"]):
        txt(ws, r, 2, b["name"], align=W.RGT, font=W.BOLD)
        txt(ws, r, 3, b["lpd"], fmt=NUM, fill=W.CALC)
        txt(ws, r, 4, b["vpd"], fmt=NUM, fill=W.CALC)
        txt(ws, r, 5, "", box=False); txt(ws, r, 6, "", box=False)
        if i < 2:
            txt(ws, r, 7, f"=C{b0+2}/C{r}-1", fmt=PCT, fill=W.CALC)
            txt(ws, r, 8, f"=D{b0+2}/D{r}-1", fmt=PCT, fill=W.CALC)
        else:
            txt(ws, r, 7, "—"); txt(ws, r, 8, "—")
        txt(ws, r, 9, "", box=False); txt(ws, r, 10, "", box=False)
        txt(ws, r, 11, ["لا أزمة — الحال اليوم كحال يناير",
                        "موسم لا خطَّ أساس — القياس عليه يصنع أزمةً ورقية",
                        "الحال اليوم"][i], align=W.WRAP,
            font=Font(name=B.FONT, size=9, color=B.INK2))
        r += 1
    for c_ in ("G", "H"): sign_cf(ws, f"{c_}{b0}:{c_}{r-1}")
    r = note(ws, r, NC,
             f"والانحدار نفسه ليس شكل منافسة: من القمة {D['brk']['peak_date']} إلى القاع "
             f"{D['brk']['trough_date']} انزلاقٌ في {D['brk']['span_days']} يوماً لا هبوطٌ "
             f"في ليلة. ورافقه هبوط حصة الديزل من {Q['1']['diesel']*100:.0f}٪ إلى "
             f"{Q['2']['diesel']*100:.0f}٪ — وحجم التعبئة يُفسَّر بالمزيج لا يُطالَب به.")
    r += 1

    r = W.band(ws, r, NC, "① ب · رقمان مختلفان عمداً — ولا تناقض بينهما")
    W.header(ws, r, ["", "الرقم", "في ورقة المحطات والحاسبة", "في هذه الورقة", "",
                     "", "", "", "", "", "لماذا يختلفان"])
    r += 1
    d0 = r
    two = [("النافذة المقيسة", "٢١٢ يوماً · يناير–يوليو",
            f"{Q['3']['days']} يوماً · الربع الثالث",
            "ورقة المحطات تستعمل نافذة الشبكة كلها ليكون القياس واحداً عبر ٥٢ محطة"),
           ("زيارة/يوم", 4260.64, Q["3"]["vpd"],
            "المحطة تعافت، فالنافذة الأحدث أعلى حركةً"),
           ("لتر لكل زيارة", 30.562, Q["3"]["lpv"],
            "وأقلّ سلّةً — لأن حصة الديزل لم تعُد إلى مستوى الربع الأول"),
           ("فجوة ساعات الذروة", 90.90, H["peak_gap"],
            "والفجوة الأحدث أصغر لأن التعافي أغلق جزءاً منها فعلاً")]
    for lab, a, b, wy in two:
        fm = NUM2 if isinstance(a, float) else None
        txt(ws, r, 2, lab, align=W.RGT, font=W.BOLD)
        txt(ws, r, 3, a, fmt=fm, fill=W.CALC)
        txt(ws, r, 4, b, fmt=fm, fill=PatternFill("solid", fgColor=B.T_GOLD),
            font=Font(name=B.FONT, size=10, bold=True, color=B.D_GOLD))
        for c_ in range(5, 11): txt(ws, r, c_, "", box=False)
        txt(ws, r, 11, wy, align=W.WRAP, font=Font(name=B.FONT, size=9, color=B.INK2))
        ws.row_dimensions[r].height = 26
        r += 1
    r = note(ws, r, NC,
             "فمن يخطّط للشبكة يستعمل ورقة المحطات — نافذتها واحدة لكل محطة. "
             "ومن يخطّط لهذه المحطة وحدها يستعمل هذه الورقة — نافذتها أحدث. "
             "والرقم الأصغر هنا ليس تحفّظاً: هو ما بقي من الفجوة بعد التعافي.",
             fill=PatternFill("solid", fgColor=B.T_ORANGE),
             font=Font(name=B.FONT, size=9, bold=True, color=B.ORANGE))
    r += 1

    # الأرباع
    r = W.band(ws, r, NC, "② الأرباع — معدّلات يومية، فالأرباع غير متساوية الطول")
    W.header(ws, r, ["", "المقياس", "الربع الأول", "الربع الثاني", "الربع الثالث",
                     "", "ر١←ر٢", "ر٢←ر٣", "ر١←ر٣", "", "القراءة"])
    r += 1
    q0 = r
    rows = [
        ("أيام مقيسة", lambda q: q["days"], NUM, "—"),
        ("معاملات/يوم", lambda q: q["vpd"], NUM, "عادت فوق مستوى الربع الأول"),
        ("لترات/يوم", lambda q: q["lpd"], NUM, "دونه قليلاً — والفارق مزيج لا حركة"),
        ("مبيعات/يوم", lambda q: q["spd"], MONEY, ""),
        ("لتر لكل معاملة", lambda q: q["lpv"], NUM2,
         "يتحرّك مع حصة الديزل لا مع سلوك العميل"),
        ("متوسط الفاتورة", lambda q: q["inv"], MONEY2, ""),
        ("حصة الديزل من اللترات", lambda q: q["diesel"], PCT,
         "المتغيّر الأكبر — وسلّة الديزل ضعف البنزين"),
        ("فواتير ٦٠ ريالاً فأكثر", lambda q: q["ge60_pct"], PCT, ""),
        ("حصة النقدي من المبيعات", lambda q: q["pay"]["cash"]["share"], PCT,
         "هبطت ١٣ نقطة في أبريل ولم تعُد — تغيّر قنوات لا تغيّر عميل"),
        ("حصة الوردية المسائية", lambda q: q["eve_share"], PCT,
         "ثابتة نحو ٦٠٪ في الأرباع الثلاثة"),
        ("زمن الخدمة", lambda q: q["dur"], '#,##0.0" ث"',
         "ارتفع في الربع الثالث مع ارتفاع الحركة"),
    ]
    for lab, fn, fm, rd in rows:
        txt(ws, r, 2, lab, align=W.RGT, font=W.BOLD)
        for i, k in enumerate(("1", "2", "3")):
            txt(ws, r, 3 + i, fn(Q[k]), fmt=fm, fill=W.CALC)
        txt(ws, r, 6, "", box=False)
        if lab == "أيام مقيسة":
            for j in range(3):
                txt(ws, r, 7 + j, "—", fill=W.CALC,
                    font=Font(name=B.FONT, size=10, color=B.INK3))
        elif fm in (PCT, PCT2):
            for j, (a, b) in enumerate(((4, 3), (5, 4), (5, 3))):
                txt(ws, r, 7 + j,
                    f"={get_column_letter(a)}{r}-{get_column_letter(b)}{r}",
                    fmt='+0.0%;-0.0%', fill=W.CALC)
        else:
            for j, (a, b) in enumerate(((4, 3), (5, 4), (5, 3))):
                txt(ws, r, 7 + j,
                    f"={get_column_letter(a)}{r}/{get_column_letter(b)}{r}-1",
                    fmt=PCT, fill=W.CALC)
        txt(ws, r, 10, "", box=False)
        txt(ws, r, 11, rd, align=W.WRAP, font=Font(name=B.FONT, size=9, color=B.INK2))
        r += 1
    for c_ in ("G", "H", "I"):
        ws.conditional_formatting.add(f"{c_}{q0}:{c_}{r-1}", CellIsRule(
            operator="lessThan", formula=["0"],
            font=Font(name=B.FONT, size=10, color=B.BAD)))
        ws.conditional_formatting.add(f"{c_}{q0}:{c_}{r-1}", CellIsRule(
            operator="greaterThan", formula=["0"],
            font=Font(name=B.FONT, size=10, color=B.D_GOOD)))
    r = note(ws, r, NC,
             "وأعمدة الفرق تُقرأ بحسب سطرها: صفوف الأعداد (معاملات · لترات · "
             "مبيعات · لتر لكل معاملة) فرقُها نسبةٌ مئوية، وصفوف النسب "
             "(ديزل · نقدي · مسائية · ٦٠ فأكثر) فرقُها نقاطٌ مئوية.",
             fill=PatternFill("solid", fgColor=B.T_NEUTRAL))
    r = note(ws, r, NC,
             "اقرأ السطرين معاً: المعاملات ر١←ر٣ موجبة واللترات سالبة. أي أن عدد "
             "العملاء لم ينقص — نقصت لترات الفاتورة، ونقصت لأن الديزل نقص. "
             "ومن يبني حملةً على «تراجع المبيعات» هنا يبيع لعميلٍ لم يذهب أصلاً.")
    r += 1

    # المنتجات
    r = W.band(ws, r, NC, "③ المنتجات — أين تحرّك المزيج بالضبط")
    W.header(ws, r, ["", "المنتج", "ر١ لتر/يوم", "ر٢ لتر/يوم", "ر٣ لتر/يوم", "",
                     "حصة ر١", "حصة ر٣", "الفرق", "", "لتر/معاملة ر٣"])
    r += 1
    p0 = r
    for p in ("g91", "g95", "diesel"):
        txt(ws, r, 2, Q["3"]["prod"][p]["name"], align=W.RGT, font=W.BOLD)
        for i, k in enumerate(("1", "2", "3")):
            txt(ws, r, 3 + i, Q[k]["prod"][p]["litres"] / Q[k]["days"], fmt=NUM, fill=W.CALC)
        txt(ws, r, 6, "", box=False)
        txt(ws, r, 7, Q["1"]["prod"][p]["share"], fmt=PCT, fill=W.CALC)
        txt(ws, r, 8, Q["3"]["prod"][p]["share"], fmt=PCT, fill=W.CALC)
        txt(ws, r, 9, f"=H{r}-G{r}", fmt='+0.0%;-0.0%', fill=W.CALC)
        txt(ws, r, 10, "", box=False)
        txt(ws, r, 11, Q["3"]["prod"][p]["lpv"], fmt=NUM2, fill=W.CALC)
        r += 1
    txt(ws, r, 2, "الإجمالي", align=W.RGT, font=W.BOLD, fill=W.FS)
    for i in range(3):
        c_ = 3 + i
        txt(ws, r, c_, f"=SUM({get_column_letter(c_)}{p0}:{get_column_letter(c_)}{r-1})",
            fmt=NUM, fill=W.FS, font=W.BOLD)
    txt(ws, r, 6, "", box=False)
    for c_ in (7, 8):
        txt(ws, r, c_, f"=SUM({get_column_letter(c_)}{p0}:{get_column_letter(c_)}{r-1})",
            fmt=PCT, fill=W.FS, font=W.BOLD)
    txt(ws, r, 9, "", fill=W.FS); txt(ws, r, 10, "", box=False)
    txt(ws, r, 11, Q["3"]["lpv"], fmt=NUM2, fill=W.FS, font=W.BOLD)
    r += 1
    r = note(ws, r, NC,
             f"سلّة الديزل {Q['3']['prod']['diesel']['lpv']:.1f} لتراً مقابل "
             f"{Q['3']['prod']['g91']['lpv']:.1f} للبنزين ٩١ — فنقطةٌ واحدة من حصة "
             "الديزل تُحرّك لتر الفاتورة أكثر مما يُحرّكه أي عرضٍ تسويقي.")
    r += 1

    # قنوات الدفع
    r = W.band(ws, r, NC, "④ قنوات الدفع — ما تغيّر في أبريل ليس عميلاً بل نظاماً")
    W.header(ws, r, ["", "القناة", "أول ظهور", "آخر ظهور", "أيام", "",
                     "حصة ر١", "حصة ر٢", "حصة ر٣", "", "القراءة"])
    r += 1
    reads = {"cash": "هبطت ١٣ نقطة ولم تعُد", "visa": "استقبلت جزءاً من الهبوط",
             "network": "قناةٌ جديدة بدأت ١١ يناير واتسعت في أبريل",
             "syarah": "قناةٌ جديدة بدأت ٨ أبريل",
             "petro": "توقّفت تقريباً في أبريل", "smartcard": "انتهت ٢٧ أبريل",
             "other": "غير معروف — وهو ضئيل هنا"}
    for ch in D["channels"]:
        txt(ws, r, 2, ch["name"], align=W.RGT, font=W.BOLD)
        txt(ws, r, 3, ch["first"], fill=W.CALC)
        txt(ws, r, 4, ch["last"], fill=W.CALC)
        txt(ws, r, 5, ch["days"], fmt=NUM, fill=W.CALC)
        txt(ws, r, 6, "", box=False)
        for i, k in enumerate(("1", "2", "3")):
            txt(ws, r, 7 + i, Q[k]["pay"][ch["key"]]["share"], fmt=PCT2, fill=W.CALC)
        txt(ws, r, 10, "", box=False)
        txt(ws, r, 11, reads.get(ch["key"], ""), align=W.WRAP,
            font=Font(name=B.FONT, size=9, color=B.INK2))
        r += 1
    r = note(ws, r, NC,
             "⛔ بوابة: ثلاث قنوات دخلت أو خرجت داخل الربع الثاني نفسه. فأي مقارنة "
             "ربعية لا توثّق هذا التغيير تنسب أثر النظام إلى البيع.",
             fill=PatternFill("solid", fgColor=B.T_GOLD),
             font=Font(name=B.FONT, size=9, bold=True, color=B.D_GOLD))
    r += 1

    # الأشهر + رسم
    r = W.band(ws, r, NC, "⑤ الأشهر — الانزلاق والتعافي في صورة واحدة")
    W.header(ws, r, ["", "الشهر", "أيام", "معاملات/يوم", "لترات/يوم", "",
                     "لتر/معاملة", "ديزل٪", "نقدي٪", "", "مبيعات/يوم"])
    r += 1
    m0 = r
    for k in sorted(M, key=int):
        m = M[k]
        txt(ws, r, 2, D["monthnames"][int(k)], align=W.RGT, font=W.BOLD)
        txt(ws, r, 3, m["days"], fmt=NUM, fill=W.CALC)
        txt(ws, r, 4, m["vpd"], fmt=NUM, fill=W.CALC)
        txt(ws, r, 5, m["lpd"], fmt=NUM, fill=W.CALC)
        txt(ws, r, 6, "", box=False)
        txt(ws, r, 7, m["lpv"], fmt=NUM2, fill=W.CALC)
        txt(ws, r, 8, m["diesel"], fmt=PCT, fill=W.CALC)
        txt(ws, r, 9, m["pay"]["cash"]["share"], fmt=PCT, fill=W.CALC)
        txt(ws, r, 10, "", box=False)
        txt(ws, r, 11, m["spd"], fmt=MONEY, fill=W.CALC)
        r += 1
    m1 = r - 1
    for c_ in ("D", "E"):
        ws.conditional_formatting.add(f"{c_}{m0}:{c_}{m1}",
                                      DataBarRule(start_type="min", end_type="max",
                                                  color=B.ORANGE, showValue=True))
    ch = LineChart(); ch.height, ch.width = 8.2, 24
    ch.title = "لترات في اليوم · وحصة الديزل"; ch.style = 2
    ch.y_axis.title = "لتر/يوم"
    ch.add_data(Reference(ws, min_col=5, min_row=m0 - 1, max_row=m1), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=m0, max_row=m1))
    ch2 = LineChart()
    ch2.add_data(Reference(ws, min_col=8, min_row=m0 - 1, max_row=m1), titles_from_data=True)
    ch2.y_axis.axId = 200; ch2.y_axis.title = "ديزل٪"; ch2.y_axis.crosses = "max"
    ch += ch2
    for s in ch.series: s.smooth = False
    ws.add_chart(ch, f"B{r+1}")
    r += 18
    r = note(ws, r, NC,
             "الخطّان يتحرّكان معاً — وهذا هو التشخيص كلّه في صورة: اللترات تتبع "
             "الديزل. ولو كان السبب منافسةً أو خدمةً لتحرّك عدد المعاملات، "
             "وهو لم يتحرّك بالقدر نفسه.")
    r += 1

    # الحملات
    r = W.band(ws, r, NC, "⑥ الحملات الثلاث — مقيسةً بثلاث نوافذ")
    W.header(ws, r, ["", "الحملة", "المقياس", "قبل", "أثناء", "بعد", "",
                     "الأثر أثناء", "", "", "الحكم"])
    r += 1
    for c_ in D["campaigns"]["rows"]:
        txt(ws, r, 2, c_["name"], align=W.RGT, font=W.BOLD)
        txt(ws, r, 3, c_["metric"], align=W.WRAP,
            font=Font(name=B.FONT, size=8, color=B.INK2))
        fm = NUM if c_["unit"] != "٪" else NUM2
        for i, v in enumerate((c_["pre"], c_["run"], c_["post"])):
            txt(ws, r, 4 + i, v, fmt=fm,
                fill=PatternFill("solid", fgColor=B.T_GOLD if i == 1 else B.T_NEUTRAL))
        txt(ws, r, 7, "", box=False)
        merge(ws, r, 8, 10, f"=E{r}/D{r}-1", fmt=PCT,
              fill=PatternFill("solid", fgColor=B.T_NEUTRAL), font=W.BOLD)
        v = c_["verdict"]
        keep = "يفوق التشتّت" in v            # أثرٌ يفوق التشتّت ويبقى بعده
        hold = "بلا بقاء" in v                # تحرَّك أثناءها ولم يُبقِ شيئاً
        tone = B.T_GOOD if keep else B.T_GOLD if hold else B.T_BAD
        col = B.D_GOOD if keep else B.D_GOLD if hold else B.D_BAD
        txt(ws, r, 11, v, align=W.WRAP,
            fill=PatternFill("solid", fgColor=tone),
            font=Font(name=B.FONT, size=9, bold=True, color=col))
        ws.row_dimensions[r].height = 30
        r += 1
    r = note(ws, r, NC,
             "الفطور: الصباح ارتفع أثناء الحملة ١٢٪ — لكنه ارتفع بعدها أكثر، فالارتفاع "
             "اتجاهٌ سابقٌ للحملة لا أثرٌ لها. وواش واي: حصة الفواتير الكبيرة قفزت "
             "٦ نقاط أثناء الأيام الأربعة ثم عادت دون مستوى ما قبلها — وهو ما تفعله "
             "حملة افتتاحٍ بالضبط: تحرّك ولا تُبقي. "
             f"وتشتّت خط الأساس الطبيعي {D['campaigns']['noise']*100:.1f}٪ فأي أثرٍ "
             "دونه لا يُحسب أثراً. والدرس: عرِّف الرقم والنافذة وخط الأساس قبل الإطلاق.",
             fill=PatternFill("solid", fgColor=B.T_GOLD),
             font=Font(name=B.FONT, size=9, bold=True, color=B.D_GOLD), h=46)
    return ws


# ═══════════════════════════════════════════════ ⑥ الساعة والوردية
def hours(wb):
    NC = 11
    ws = sheet(wb, S_H, [3, 12, 12, 12, 12, 12, 4, 11, 11, 9, 30],
               "الساعة والوردية — أين الفجوة بالضبط، وفي أي وردية",
               f"العمرة بالساعة (المحطة الوحيدة) ثم {SG['n']} محطة بالوردية · "
               f"المقارنة بشكل تصنيف «{D['seg']}» — وسيط {H['ref_n']} محطة", NC)
    r = 4
    r = W.band(ws, r, NC, "① العمرة — الفجوة لكل ساعة")
    W.header(ws, r, ["", "الساعة", "فعلي/يوم", "حصتنا", "شكل التصنيف",
                     "المتوقَّع/يوم", "", "الفجوة", "ذروة؟", "الوردية", "ملاحظة"])
    r += 1
    h0 = r
    for row in H["rows"]:
        h = row["hour"]
        mark = row["gap"] > 5
        fill = PatternFill("solid", fgColor=B.T_BAD if mark else B.T_NEUTRAL)
        txt(ws, r, 2, f"{h:02d}:٠٠", font=W.BOLD)
        txt(ws, r, 3, row["actual"], fmt=NUM, fill=W.CALC)
        txt(ws, r, 4, row["share"], fmt=PCT2, fill=W.CALC)
        txt(ws, r, 5, row["ref"], fmt=PCT2, fill=W.CALC)
        txt(ws, r, 6, f"=E{r}*{Q['3']['vpd']:.4f}", fmt=NUM, fill=W.CALC)
        txt(ws, r, 7, "", box=False)
        txt(ws, r, 8, f"=MAX(0,F{r}-C{r})", fmt=NUM, fill=fill,
            font=Font(name=B.FONT, size=10, bold=mark,
                      color=B.D_BAD if mark else B.INK))
        txt(ws, r, 9, "★" if row["peak"] else "",
            font=Font(name=B.FONT, size=10, color=B.ORANGE))
        txt(ws, r, 10, row["shift"], font=Font(name=B.FONT, size=9, color=B.INK2))
        txt(ws, r, 11, "ساعة تسليم الورديات" if h == 23 else "",
            align=W.WRAP, font=Font(name=B.FONT, size=8, color=B.D_BAD))
        r += 1
    h1 = r - 1
    txt(ws, r, 2, "المجموع", font=W.BOLD, fill=W.FS)
    txt(ws, r, 3, f"=SUM(C{h0}:C{h1})", fmt=NUM, fill=W.FS, font=W.BOLD)
    for c_ in (4, 5):
        txt(ws, r, c_, f"=SUM({get_column_letter(c_)}{h0}:{get_column_letter(c_)}{h1})",
            fmt=PCT, fill=W.FS, font=W.BOLD)
    txt(ws, r, 6, f"=SUM(F{h0}:F{h1})", fmt=NUM, fill=W.FS, font=W.BOLD)
    txt(ws, r, 7, "", box=False)
    txt(ws, r, 8, f"=SUM(H{h0}:H{h1})", fmt=NUM, fill=W.FS, font=W.BOLD)
    for c_ in (9, 10, 11): txt(ws, r, c_, "", fill=W.FS)
    r += 1
    ws.conditional_formatting.add(f"C{h0}:C{h1}",
                                  DataBarRule(start_type="num", start_value=0,
                                              end_type="max", color=B.GOLD, showValue=True))
    ch = BarChart(); ch.type = "col"; ch.height, ch.width = 8.5, 26
    ch.title = "المعاملات في الساعة — فعلي مقابل شكل التصنيف"
    ch.add_data(Reference(ws, min_col=3, min_row=h0 - 1, max_row=h1), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=h0, max_row=h1))
    ln = LineChart()
    ln.add_data(Reference(ws, min_col=6, min_row=h0 - 1, max_row=h1), titles_from_data=True)
    ch += ln
    ws.add_chart(ch, f"B{r}")
    r += 18

    ho = D["handover"]
    r = note(ws, r, NC,
             f"الساعة {ho['row']['hour']:02d}:٠٠ فيها {ho['row']['gap']:.0f} معاملة/يوم "
             f"فاقداً — أكبر من أي ساعة ذروة. وتسقط من الرقم المحافظ "
             f"({H['peak_gap']:.0f} معاملة/يوم) لأنها ليست ضمن ساعاتنا الثماني الأعلى، "
             f"وليست كذلك لأننا ننخفض فيها. و{ho['note']}. "
             "فإغلاقها يُختبر أولاً بتغيير موعد التسليم — لا بحملة.",
             fill=PatternFill("solid", fgColor=B.T_ORANGE),
             font=Font(name=B.FONT, size=9, bold=True, color=B.ORANGE), h=42)
    r += 1

    # وردية العمرة
    r = W.band(ws, r, NC, "② العمرة — الوردية: الفجوة كلّها في جهةٍ واحدة")
    W.header(ws, r, ["", "المقياس", "الصباحية ٠٠–١١", "المسائية ١٢–٢٣", "الإجمالي",
                     "", "", "", "", "", "القراءة"])
    r += 1
    s0 = r
    srows = [
        ("معاملات/يوم", SH["morning"], SH["evening"], NUM, "المسائية ٦٠٪ من الطلب"),
        ("عدد العمّال", SH["day_w"], SH["eve_w"], NUM,
         "وهي ٤٥٪ من العمالة — الطلب في جهةٍ والطاقم في أخرى"),
        ("معاملات لكل عامل", None, None, NUM,
         f"العامل المسائي يخدم {SH['ratio']:.2f}× ما يخدمه الصباحي"),
        ("الفجوة في ساعات الذروة", H["peak_m"], H["peak_e"], NUM2,
         "الرقم المؤكَّد — وكلّه في المسائية"),
        ("الفجوة في كل الساعات", H["gap_m"], H["gap_e"], NUM2,
         "يشمل ساعة تسليم الورديات وساعات الهدوء — لا يُطالَب به"),
        ("العدد المتوازن للعمّال", SH["need_m"], SH["need_e"], NUM2,
         f"أي نقل {abs(SH['move']):.0f} عمّال من الصباحية إلى المسائية — بلا توظيف"),
    ]
    for lab, a, b, fm, rd in srows:
        txt(ws, r, 2, lab, align=W.RGT, font=W.BOLD)
        if lab == "معاملات لكل عامل":
            txt(ws, r, 3, f"=C{s0}/C{s0+1}", fmt=fm,
                fill=PatternFill("solid", fgColor=B.T_GOOD),
                font=Font(name=B.FONT, size=11, bold=True, color=B.D_GOOD))
            txt(ws, r, 4, f"=D{s0}/D{s0+1}", fmt=fm,
                fill=PatternFill("solid", fgColor=B.T_BAD),
                font=Font(name=B.FONT, size=11, bold=True, color=B.D_BAD))
        else:
            txt(ws, r, 3, a, fmt=fm, fill=W.CALC)
            txt(ws, r, 4, b, fmt=fm, fill=W.CALC)
        txt(ws, r, 5,
            f"=E{s0}/E{s0+1}" if lab == "معاملات لكل عامل" else f"=C{r}+D{r}",
            fmt=fm, fill=W.FS, font=W.BOLD)
        for c_ in range(6, 11): txt(ws, r, c_, "", box=False)
        txt(ws, r, 11, rd, align=W.WRAP, font=Font(name=B.FONT, size=9, color=B.INK2))
        r += 1
    r = note(ws, r, NC,
             f"وسعة الخدمة تؤكّد القراءة: زمن الخدمة المقيس {SH['dur']:.0f} ثانية، "
             f"فالعامل الواحد يخدم {SH['cap_worker']:.0f} معاملة في الساعة نظرياً، "
             f"و{SH['eve_w']:.0f} عمّال يخدمون {SH['peak_cap']:.0f}. وساعة الذروة "
             f"({SH['peak_hour']:02d}:٠٠) فيها {SH['peak_txn']:.0f} معاملة — أي "
             f"{SH['utilisation']*100:.0f}٪ من السعة. "
             "والمحطة عند هذا الحدّ لا تُزاد مبيعاتها بعرضٍ بل بيدٍ إضافية أو بخدمةٍ أسرع.",
             fill=PatternFill("solid", fgColor=B.T_GOLD),
             font=Font(name=B.FONT, size=9, bold=True, color=B.D_GOLD), h=34)
    r += 1

    # الشبكة بالوردية
    r = W.band(ws, r, NC,
               f"③ الشبكة — {SG['n']} محطة عليها بيان وردية · اختلال الحِمل وقابلية النقل")
    W.header(ws, r, ["", "المحطة", "التصنيف", "عمّال", "حصة الطلب المسائية",
                     "حصة العمالة المسائية", "", "الاختلال", "حِمل صباحي",
                     "حِمل مسائي", "عمّال قابلون للنقل"])
    r += 1
    g0 = r
    rows = sorted(SG["rows"], key=lambda x: -x["gap"])
    for s in rows:
        txt(ws, r, 2, s["name"], align=W.RGT, font=W.BOLD)
        txt(ws, r, 3, s["seg"], font=Font(name=B.FONT, size=9, color=B.INK2))
        txt(ws, r, 4, s["staff"], fmt=NUM, fill=W.CALC)
        txt(ws, r, 5, s["eve"], fmt=PCT, fill=W.CALC)
        txt(ws, r, 6, s["eve_staff"], fmt=PCT, fill=W.CALC)
        txt(ws, r, 7, "", box=False)
        txt(ws, r, 8, f"=E{r}-F{r}", fmt='+0.0%;-0.0%', fill=W.CALC, font=W.BOLD)
        txt(ws, r, 9, s["load_m"], fmt=NUM, fill=W.CALC)
        txt(ws, r, 10, s["load_e"], fmt=NUM, fill=W.CALC)
        txt(ws, r, 11, s["move"], fmt=NUM2, fill=W.CALC)
        r += 1
    g1 = r - 1
    txt(ws, r, 2, "الإجمالي", align=W.RGT, font=W.BOLD, fill=W.FS)
    txt(ws, r, 3, "", fill=W.FS)
    txt(ws, r, 4, f"=SUM(D{g0}:D{g1})", fmt=NUM, fill=W.FS, font=W.BOLD)
    txt(ws, r, 5, SG["eve_txn"], fmt=PCT, fill=W.FS, font=W.BOLD)
    txt(ws, r, 6, "", fill=W.FS); txt(ws, r, 7, "", box=False)
    for c_ in (8, 9, 10): txt(ws, r, c_, "", fill=W.FS)
    txt(ws, r, 11, f"=SUM(K{g0}:K{g1})", fmt=NUM2, fill=W.FS, font=W.BOLD)
    r += 1
    sign_cf(ws, f"H{g0}:H{g1}", good_high=False)
    ws.conditional_formatting.add(f"J{g0}:J{g1}",
                                  DataBarRule(start_type="num", start_value=0,
                                              end_type="max", color=B.BAD, showValue=True))
    ws.auto_filter.ref = f"B{g0-1}:K{g1}"
    r = note(ws, r, NC,
             f"«الاختلال» = حصة الطلب المسائية ناقص حصة العمالة المسائية. وهو موجبٌ في "
             f"أغلب الشبكة: المساء {SG['eve_txn']*100:.0f}٪ من المعاملات "
             f"و{SG['eve_rev']*100:.0f}٪ من الإيراد. "
             f"ومجموع العمّال القابلين للنقل {SG['movable']:.0f} عاملاً عبر "
             f"{SG['n']} محطة — نقلٌ لا توظيف، وكلفته صفر. "
             f"و{SG['no_worker']} محطة بلا بيان عامل أصلاً فلا تدخل الحساب.",
             fill=PatternFill("solid", fgColor=B.T_GOLD),
             font=Font(name=B.FONT, size=9, bold=True, color=B.D_GOLD), h=40)
    return ws


# ═══════════════════════════════════════════════ ⑦ كلفة اللتر والمستهدف
def costsheet(wb):
    NC = 10
    ws = sheet(wb, S_C, [3, 28, 13, 13, 13, 13, 13, 4, 13, 30],
               "كلفة اللتر — لا قيمته",
               "القيمة تُقاس بالهامش، والتكلفة بما نُخرجه من جيبنا لنبيع لتراً: "
               "شراؤه ثم تشغيل المحطة التي باعته ثم ما يُحمَّل عليه من المركز · "
               f"المصدر: قائمة الدخل · {LC['pl']['total_stations']} محطة · "
               f"{LC['pl']['total_litres']/1e6:,.0f} مليون لتر", NC)
    r = 4
    r = W.band(ws, r, NC, "① الطبقات الأربع — هللة لكل لتر، وبتفصيل نموذج العمل")
    W.header(ws, r, ["", "الطبقة", "الشبكة"] + LC["pl"]["models"] + ["", "ألف ريال", "ما تعنيه"])
    r += 1
    c0 = r
    for st in LC["pl"]["stack"]:
        txt(ws, r, 2, st["layer"], align=W.RGT, font=W.BOLD)
        txt(ws, r, 3, st["cpl"], fmt=HAL, fill=W.CALC, font=W.BOLD)
        for i, v in enumerate(st["by_model"]):
            txt(ws, r, 4 + i, v, fmt=HAL, fill=W.CALC)
        txt(ws, r, 8, "", box=False)
        txt(ws, r, 9, st["total"], fmt=MONEY, fill=W.CALC)
        txt(ws, r, 10, st["note"], align=W.WRAP,
            font=Font(name=B.FONT, size=9, color=B.INK2))
        ws.row_dimensions[r].height = 26
        r += 1
    c1 = r - 1
    txt(ws, r, 2, "مجموع الكلفة", align=W.RGT, font=W.BOLD, fill=W.FS)
    for c_ in range(3, 8):
        txt(ws, r, c_, f"=SUM({get_column_letter(c_)}{c0}:{get_column_letter(c_)}{c1})",
            fmt=HAL, fill=W.FS, font=W.BOLD)
    txt(ws, r, 8, "", box=False)
    txt(ws, r, 9, f"=SUM(I{c0}:I{c1})", fmt=MONEY, fill=W.FS, font=W.BOLD)
    txt(ws, r, 10, "", fill=W.FS)
    r += 1
    r = note(ws, r, NC,
             "٩٧٪ من كلفة اللتر شراؤه من أرامكو — ولا قرار لنا فيه. "
             "والقرار كلّه في الطبقة الثانية: تشغيل المحطة. "
             "ولذلك لا يُقاس أداء التجاري على الطبقة الأولى ولا يُلام عليها. "
             "⚠ ولا يُطرح هذا الجدول من جدول ② أدناه: هامش المساهمة هناك محسوبٌ "
             "بعد الطبقة الأولى أصلاً، فطرحُها ثانيةً يحتسبها مرتين.",
             fill=PatternFill("solid", fgColor=B.T_GOLD),
             font=Font(name=B.FONT, size=9, bold=True, color=B.D_GOLD))
    r += 1

    # قيمة اللتر بالمنتج
    r = W.band(ws, r, NC, "② قيمة اللتر بالمنتج — الهامش ليس واحداً")
    W.header(ws, r, ["", "المنتج", "سعر المضخة", "قبل الضريبة", "هامش اللتر",
                     "خصم التشغيل", "الصافي", "", "الوزن في المزيج", "الملاحظة"])
    r += 1
    l0 = r
    for row in LIT["rows"]:
        txt(ws, r, 2, row["fuel"], align=W.RGT, font=W.BOLD)
        txt(ws, r, 3, row["price"], fmt=MONEY2, fill=W.CALC)
        txt(ws, r, 4, row["net_price"], fmt=MONEY2, fill=W.CALC)
        txt(ws, r, 5, row["cpl"], fmt=HAL, fill=W.CALC, font=W.BOLD)
        txt(ws, r, 6, -K["opex"] * 100, fmt=HAL, fill=W.CALC)
        txt(ws, r, 7, f"=E{r}+F{r}", fmt=HAL, fill=W.CALC, font=W.BOLD)
        txt(ws, r, 8, "", box=False)
        txt(ws, r, 9, row["w"], fmt=PCT, fill=W.CALC)
        txt(ws, r, 10, "من انحدار ١٩ محطة — لا من جدول معتمد", align=W.WRAP,
            font=Font(name=B.FONT, size=8, color=B.INK3))
        r += 1
    l1 = r - 1
    txt(ws, r, 2, "المرجَّح بالمزيج", align=W.RGT, font=W.BOLD, fill=W.FS)
    for c_ in (3, 4): txt(ws, r, c_, "", fill=W.FS)
    txt(ws, r, 5, f"=SUMPRODUCT(E{l0}:E{l1},$I${l0}:$I${l1})", fmt=HAL,
        fill=W.FS, font=W.BOLD)
    txt(ws, r, 6, -K["opex"] * 100, fmt=HAL, fill=W.FS)
    txt(ws, r, 7, f"=E{r}+F{r}", fmt=HAL, fill=W.FS, font=W.BOLD)
    txt(ws, r, 8, "", box=False)
    txt(ws, r, 9, f"=SUM(I{l0}:I{l1})", fmt=PCT, fill=W.FS, font=W.BOLD)
    txt(ws, r, 10, "", fill=W.FS)
    sign_cf(ws, f"G{l0}:G{r}")
    r += 1
    r = note(ws, r, NC,
             f"⛔ تحفُّظ صريح: هامشا البنزين ({K['margin_petrol']*100:.2f} هللة) والديزل "
             f"({K['margin_diesel']*100:.2f}) مستنتجان بانحدار هامش ١٩ محطة على حصة "
             "ديزلها، وأقصى حصة ديزل مرصودة ٤٣٪ — فرقم الديزل استقراءٌ خارج المدى "
             "لا قياس. ويسنده هيكل عمولتنا (٣ هللات بنزين · ١٫٥ ديزل) ولا يغني عن "
             "جدول هامش معتمد من المالية.",
             fill=PatternFill("solid", fgColor=B.T_GOLD),
             font=Font(name=B.FONT, size=9, bold=True, color=B.D_GOLD), h=36)
    r += 1

    # داخل مصاريف التشغيل
    r = W.band(ws, r, NC, "③ داخل «تشغيل المحطة» — أجورٌ ومرافق، لأعلى وأدنى المحطات")
    W.header(ws, r, ["", "المحطة", "الرمز", "لتر الربع", "عمّال", "أجر/لتر",
                     "مرافق/لتر", "", "لتر لكل عامل", "القراءة"])
    r += 1
    o = LC["op"]["rows"]
    show = o[:5] + o[-3:]
    o0 = r
    for i, s in enumerate(show):
        sep = i == 5
        txt(ws, r, 2, s["name"], align=W.RGT, font=W.BOLD)
        txt(ws, r, 3, s["code"], font=Font(name=B.FONT, size=9, color=B.INK2))
        txt(ws, r, 4, s["vol"], fmt=NUM, fill=W.CALC)
        txt(ws, r, 5, s["staff"], fmt=NUM, fill=W.CALC)
        txt(ws, r, 6, s["wage_cpl"], fmt=HAL, fill=W.CALC, font=W.BOLD)
        txt(ws, r, 7, s["util_cpl"], fmt=HAL, fill=W.CALC)
        txt(ws, r, 8, "", box=False)
        txt(ws, r, 9, s["litres_per_worker"], fmt=NUM, fill=W.CALC)
        txt(ws, r, 10, "أعلى خمس كلفةَ عمالة" if i < 5 else "أدناها", align=W.WRAP,
            font=Font(name=B.FONT, size=8,
                      color=B.D_BAD if i < 5 else B.D_GOOD))
        r += 1
    ws.conditional_formatting.add(f"F{o0}:F{r-1}",
                                  DataBarRule(start_type="num", start_value=0,
                                              end_type="max", color=B.BAD, showValue=True))
    r = note(ws, r, NC,
             "⛔ ومعدّل الأجر في هذا التقرير ليس مقيساً: الأجور ÷ عدد العمّال = "
             "٣٬٥٠٠ ريالاً بالضبط في ٢١ محطة بلا استثناء واحد — أي أنه سعرٌ ثابت "
             "يُضرب في العدد، وهو أجرٌ شهري طُبِّق مرة بدل ثلاث. "
             "ولذلك مصاريف التقرير ٤٫٨٦ هللة بينما قائمة الدخل تقول ٩٫١٥ للمحطات "
             "نفسها — ٢٫٠٢×. فالقائمة هي المرجع، والتقرير للمقارنة النسبية بين "
             "المحطات لا لمستوى الكلفة المطلق.",
             fill=PatternFill("solid", fgColor=B.T_GOLD),
             font=Font(name=B.FONT, size=9, bold=True, color=B.D_GOLD), h=46)
    r += 1

    # ما لم يُقس
    r = W.band(ws, r, NC, "④ ما لم يدخل الحساب لأنه غير مقيس — ولا يُملأ بتقدير")
    W.header(ws, r, ["", "البند", "الحالة", "", "", "", "لماذا يهمّ", "", "", "المالك"])
    r += 1
    for name, state, why, owner in C["unmeasured"]:
        txt(ws, r, 2, name, align=W.RGT, font=W.BOLD)
        merge(ws, r, 3, 6, state, fill=PatternFill("solid", fgColor=B.T_GOLD),
              font=Font(name=B.FONT, size=10, bold=True, color=B.D_GOLD))
        merge(ws, r, 7, 9, why, align=W.RGT,
              font=Font(name=B.FONT, size=9, color=B.INK2))
        txt(ws, r, 10, owner, align=W.WRAP,
            font=Font(name=B.FONT, size=9, color=B.INK2))
        ws.row_dimensions[r].height = 28
        r += 1
    r = note(ws, r, NC,
             "وهذه البنود تُقتطع من هامش كل لتر إضافي، فكلّما بقيت غير مقيسة بقي "
             "الصافي أعلى مما هو. والفرق بينها وبين التقدير أنها مكتوبة هنا باسمها "
             "وباسم من يملك سدّها.")
    return ws


# ═══════════════════════════════════════════════ ⑧ خطة المستهدف
def plan(wb):
    NC = 11
    ws = sheet(wb, S_T, [3, 26, 12, 12, 12, 12, 4, 13, 13, 4, 30],
               "خطة المستهدف — بالمعاملات لا باللترات",
               "المطلوب من الوردية عددُ عملاء تخدمهم لا لترات يضخّونها: "
               "اللترات تتبع المزيج ولا يملكها العامل", NC)
    r = 4
    r = note(ws, r, NC,
             "⛔ قبل أي مستهدف: البوابات تُقفل أولاً. والمستهدف أدناه صالحٌ للحساب "
             "والعرض، ولا يُحمَّل على وردية قبل إقرار جدول الورديات من العمليات.",
             fill=PatternFill("solid", fgColor=B.T_GOLD),
             font=Font(name=B.FONT, size=10, bold=True, color=B.D_GOLD))
    r += 1

    r = W.band(ws, r, NC, "① الشبكة — المستهدف بالتصنيف")
    W.header(ws, r, ["", "التصنيف", "محطات", "دون المتوقَّع", "فجوة الذروة",
                     "فاقد السلة لتر/يوم", "", "قيمة المعاملات", "قيمة السلة",
                     "", "المحرّك الأول والفعل"])
    r += 1
    g0 = r
    for s in SEGS:
        txt(ws, r, 2, s["seg"], align=W.RGT, font=W.BOLD)
        txt(ws, r, 3, s["n"], fmt=NUM, fill=W.CALC)
        txt(ws, r, 4, s["below"], fmt=NUM, fill=W.CALC)
        txt(ws, r, 5, s["gap_peak"], fmt=NUM, fill=W.CALC)
        txt(ws, r, 6, s["upl_fill"], fmt=NUM, fill=W.CALC)
        txt(ws, r, 7, "", box=False)
        txt(ws, r, 8, s["sar_txn"], fmt=MONEY, fill=W.CALC)
        txt(ws, r, 9, s["sar_fill"], fmt=MONEY, fill=W.CALC)
        txt(ws, r, 10, "", box=False)
        txt(ws, r, 11, f"{s['driver']} — {s['action']}", align=W.WRAP,
            font=Font(name=B.FONT, size=9, color=B.INK2))
        ws.row_dimensions[r].height = 26
        r += 1
    g1 = r - 1
    txt(ws, r, 2, "الإجمالي", align=W.RGT, font=W.BOLD, fill=W.FS)
    for c_ in (3, 4, 5, 6):
        txt(ws, r, c_, f"=SUM({get_column_letter(c_)}{g0}:{get_column_letter(c_)}{g1})",
            fmt=NUM, fill=W.FS, font=W.BOLD)
    txt(ws, r, 7, "", box=False)
    for c_ in (8, 9):
        txt(ws, r, c_, f"=SUM({get_column_letter(c_)}{g0}:{get_column_letter(c_)}{g1})",
            fmt=MONEY, fill=W.FS, font=W.BOLD)
    txt(ws, r, 10, "", box=False); txt(ws, r, 11, "", fill=W.FS)
    r += 1
    r = note(ws, r, NC,
             f"وفاقد ساعات الهدوء {TOT['gap_quiet']:,.0f} معاملة/يوم غير مُدرَج عمداً: "
             "النقص في ساعات الهدوء اختلاف طلبٍ لا فاقد، ولا يُطالَب به من لا يملكه.")
    r += 1

    r = W.band(ws, r, NC, "② أعلى خمس عشرة محطة قيمةً — ومن يملك كلاً منها")
    W.header(ws, r, ["", "المحطة", "التصنيف", "فجوة الذروة", "فجوة السلة",
                     "الإجمالي ر/سنة", "", "المطلوب/يوم", "قيمة الإغلاق", "",
                     "الفعل والبوابة"])
    r += 1
    t0 = r
    for s in ST[:15]:
        gate = any(c_ in s["conds"] for c_ in
                   ("فجوة مطابقة", "تغيير مسار", "بيانات دفع مفقودة"))
        fill = PatternFill("solid", fgColor=B.T_GOLD) if gate else None
        txt(ws, r, 2, s["name"], align=W.RGT, font=W.BOLD, fill=fill)
        txt(ws, r, 3, s["seg"], font=Font(name=B.FONT, size=9, color=B.INK2), fill=fill)
        txt(ws, r, 4, s["gap_peak"], fmt=NUM2, fill=fill or W.CALC)
        txt(ws, r, 5, s["gap_fill"], fmt=NUM2, fill=fill or W.CALC)
        txt(ws, r, 6, s["sar_total"], fmt=MONEY, fill=fill or W.CALC, font=W.BOLD)
        txt(ws, r, 7, "", box=False)
        txt(ws, r, 8, f"=D{r}*{K['close']}", fmt=NUM2, fill=fill or W.CALC)
        txt(ws, r, 9, f"=F{r}*{K['close']}", fmt=MONEY, fill=fill or W.CALC)
        txt(ws, r, 10, "", box=False)
        act = ("⛔ " + " · ".join(s["conds"])) if gate else \
            next((sg["action"] for sg in SEGS if sg["seg"] == s["seg"]), "")
        txt(ws, r, 11, act, align=W.WRAP, fill=fill,
            font=Font(name=B.FONT, size=8,
                      color=B.D_GOLD if gate else B.INK2))
        r += 1
    t1 = r - 1
    txt(ws, r, 2, "مجموع الخمس عشرة", align=W.RGT, font=W.BOLD, fill=W.FS)
    txt(ws, r, 3, "", fill=W.FS)
    for c_ in (4, 5): txt(ws, r, c_, f"=SUM({get_column_letter(c_)}{t0}:{get_column_letter(c_)}{t1})",
                          fmt=NUM2, fill=W.FS, font=W.BOLD)
    txt(ws, r, 6, f"=SUM(F{t0}:F{t1})", fmt=MONEY, fill=W.FS, font=W.BOLD)
    txt(ws, r, 7, "", box=False)
    txt(ws, r, 8, f"=SUM(H{t0}:H{t1})", fmt=NUM2, fill=W.FS, font=W.BOLD)
    txt(ws, r, 9, f"=SUM(I{t0}:I{t1})", fmt=MONEY, fill=W.FS, font=W.BOLD)
    txt(ws, r, 10, "", box=False); txt(ws, r, 11, "", fill=W.FS)
    r += 1
    ws.conditional_formatting.add(f"F{t0}:F{t1}",
                                  DataBarRule(start_type="num", start_value=0,
                                              end_type="max", color=B.ORANGE, showValue=True))
    r += 1

    # القياس
    r = W.band(ws, r, NC, "③ القياس قبل التنفيذ — بأي رقم نحكم ومتى")
    W.header(ws, r, ["", "السؤال", "الجواب", "", "", "", "", "", "", "", "لماذا"])
    r += 1
    qs = [
        ("ما الرقم الذي نحكم به؟",
         "معاملات ساعات ذروة المحطة وحدها — لا مبيعاتها كلها",
         "مبيعات المحطة تتحرّك بالمزيج والموسم، وهما خارج يد الوردية"),
        ("مقابل أي خط أساس؟",
         "الأسابيع الأربعة السابقة للتنفيذ في المحطة نفسها",
         "المقارنة بالعام الماضي تخلط الموسم بالأثر"),
        ("متى نحكم؟", "بعد ستة أسابيع — ثم كل شهر",
         "أقصر من ذلك يقرأ التشتّت أثراً"),
        ("ما حدّ الأثر المقبول؟",
         f"أكثر من {D['campaigns']['noise']*100:.1f}٪ — وهو تشتّت خط الأساس المقيس",
         "ما دونه لا يُحسب أثراً مهما بدا موجباً"),
        ("من يملك الرقم؟", "مدير المحطة للتنفيذ · التجاري للقياس · العمليات للطاقم",
         "الهدف بلا مالكٍ واحد لا يُقاس ولا يُحاسَب عليه"),
        ("ما الذي لا نطالب به؟", "لترات الفاتورة الواحدة",
         "ثلثا تباينها مزيج وقود لا سلوك — تُفسَّر ولا تُطلَب"),
    ]
    for q_, a, why in qs:
        txt(ws, r, 2, q_, align=W.RGT, font=W.BOLD)
        merge(ws, r, 3, 10, a, align=W.RGT,
              font=Font(name=B.FONT, size=10, bold=True, color=B.ORANGE))
        txt(ws, r, 11, why, align=W.WRAP, font=Font(name=B.FONT, size=9, color=B.INK2))
        ws.row_dimensions[r].height = 26
        r += 1
    r = note(ws, r, NC,
             "ما لا يُصمَّم قياسه قبلاً تُنسب نتيجته بعداً إلى أي سبب. "
             "والسؤال الكاشف لأي خطة: ما آخر هدفٍ بيعيّ رفضتَه، ولماذا؟",
             fill=PatternFill("solid", fgColor=B.T_ORANGE),
             font=Font(name=B.FONT, size=9, bold=True, color=B.ORANGE))
    return ws


# ═══════════════════════════════════════════════ ⑨ المنافسون
def competitors(wb):
    NC = 10
    R_ = X["retail"]; CO = X["corporate"]; PR = X["property"]
    ws = sheet(wb, "المنافسون بالمنتج", [3, 24, 13, 13, 13, 13, 4, 13, 4, 32],
               "لكل منتج منافسٌ مختلف ومقياسٌ مختلف ومصدرُ بيانات مختلف",
               "ولا يُجمعون في جدول واحد — فمن يقارن وقود الأفراد بالعقار "
               "يقارن سوقين لا منافسين", NC)
    r = 4
    r = W.band(ws, r, NC, "① وقود الأفراد — المسح الميداني حول محطاتنا")
    W.header(ws, r, ["", "العلامة", "مواقع", "الحصة", "أقرب مسافة م",
                     "وسيط المسافة م", "", "التقييم", "", "الملاحظة"])
    r += 1
    b0 = r
    for b in R_["brands"]:
        txt(ws, r, 2, b["brand"], align=W.RGT, font=W.BOLD)
        txt(ws, r, 3, b["n"], fmt=NUM, fill=W.CALC)
        txt(ws, r, 4, b["share"], fmt=PCT, fill=W.CALC)
        txt(ws, r, 5, b["near"], fmt=NUM, fill=W.CALC)
        txt(ws, r, 6, b["med"], fmt=NUM, fill=W.CALC)
        txt(ws, r, 7, "", box=False)
        txt(ws, r, 8, b["rating"], fmt='0.00"★"', fill=W.CALC)
        txt(ws, r, 9, "", box=False)
        txt(ws, r, 10, f"{b['reviews']:,.0f} مراجعة — وسيط", align=W.WRAP,
            font=Font(name=B.FONT, size=8, color=B.INK3))
        r += 1
    txt(ws, r, 2, "درب", align=W.RGT, font=W.BOLD, fill=W.SECT)
    txt(ws, r, 3, R_["n_total"], fmt=NUM, fill=W.SECT, font=W.BOLD)
    for c_ in (4, 5, 6): txt(ws, r, c_, "", fill=W.SECT)
    txt(ws, r, 7, "", box=False)
    txt(ws, r, 8, R_["ours_mean"], fmt='0.00"★"', fill=W.SECT,
        font=Font(name=B.FONT, size=11, bold=True, color=B.D_GOOD))
    txt(ws, r, 9, "", box=False)
    txt(ws, r, 10, f"فارقنا {R_['gap_mean']:+.2f} نجمة", align=W.WRAP,
        font=Font(name=B.FONT, size=9, bold=True, color=B.D_GOOD), fill=W.SECT)
    r += 1
    r = note(ws, r, NC,
             f"⛔ وهنا الفحص الثاني: تقييمنا {R_['ours_mean']:.2f}★ مقابل "
             f"{R_['rate_mean']:.2f} للمنافسين، وهو ثابتٌ على المواقع كلها. "
             f"والنمو فيها يمتدّ من {R_['growth_hi']:+.0f}٪ إلى "
             f"{R_['growth_lo']:+.0f}٪. "
             "وما يستقرّ لا يفسّر ما يتباين — فأياً كان محرّك النتيجة فليس هو "
             f"التفوّق الخدمي. وارتباط فجوة التقييم بالنمو "
             f"{R_['r_gap']:+.2f} على {R_['n_sample']} موقعاً، "
             f"وارتباط عدد المنافسين بالنمو {R_['r_n']:+.2f}: لا شيء في كليهما.",
             fill=PatternFill("solid", fgColor=B.T_GOLD),
             font=Font(name=B.FONT, size=9, bold=True, color=B.D_GOLD), h=44)
    r += 1

    r = W.band(ws, r, NC, "② وقود الشركات — والتصنيف بحجم السلة لا بالاسم")
    W.header(ws, r, ["", "القناة", "الإيراد ريال", "الزيارات", "متوسط الفاتورة",
                     "الحصة", "", "التصنيف", "", "الدليل"])
    r += 1
    for p in CO["pays"]:
        txt(ws, r, 2, p.get("ar") or p.get("pay", ""), align=W.RGT, font=W.BOLD)
        txt(ws, r, 3, p["rev"], fmt=MONEY, fill=W.CALC)
        txt(ws, r, 4, p["vis"], fmt=NUM, fill=W.CALC)
        txt(ws, r, 5, f"=C{r}/D{r}", fmt=MONEY2, fill=W.CALC, font=W.BOLD)
        txt(ws, r, 6, p["share"], fmt=PCT, fill=W.CALC)
        txt(ws, r, 7, "", box=False)
        fleet = bool(p.get("plat")) or bool(p.get("fleet_basket"))
        kind = ("أسطول" if fleet else "غير مسجَّلة"
                if not p.get("ar") or "غير" in str(p.get("ar")) else "نقاط بيع")
        why = ("سلّتها ثلاثة أضعاف النقدي — قناة أسطول"
               if fleet else "وسيلة الدفع غير مسجَّلة — بوابة بيانات"
               if kind == "غير مسجَّلة" else
               "سلّتها بحجم النقدي — جهاز نقاط بيع لا أسطول")
        tone = (B.T_GOOD if fleet else B.T_GOLD if kind == "غير مسجَّلة"
                else B.T_NEUTRAL)
        col = (B.D_GOOD if fleet else B.D_GOLD if kind == "غير مسجَّلة" else B.INK)
        txt(ws, r, 8, kind, fill=PatternFill("solid", fgColor=tone),
            font=Font(name=B.FONT, size=9, bold=True, color=col))
        txt(ws, r, 9, "", box=False)
        txt(ws, r, 10, why, align=W.WRAP,
            font=Font(name=B.FONT, size=8, color=B.INK3))
        r += 1
    r = note(ws, r, NC,
             f"«الجهاز الخارجي» اسمٌ يوحي بقناة أسطول، وفاتورته "
             f"{CO['pos_inv']:.2f} ريالاً — بحجم النقدي {CO['cash_inv']:.2f}. "
             "فهو جهاز نقاط بيع. وحين حُسب أسطولاً انتفخت حصة الأسطول في كورنيش "
             "الخبر من ٠٫٩٪ الحقيقية إلى ٥١٫٣٪، فانتفخت معها شريحةٌ كاملة.")
    r += 1

    r = W.band(ws, r, NC, "③ العقار والتأجير — بديل المستأجر لا متر المربّع")
    W.header(ws, r, ["", "الصيغة", "الوحدات", "المؤجَّرة", "الشاغرة", "الإشغال",
                     "", "", "", "بديل المستأجر"])
    r += 1
    p0 = r
    shop = next((t for t in PR["types"] if t["key"] == "shop"), None)
    for t in sorted(PR["types"], key=lambda t: -t["occ"]):
        low = t["occ"] < 0.30
        txt(ws, r, 2, t["ar"], align=W.RGT, font=W.BOLD)
        txt(ws, r, 3, t["n"], fmt=NUM, fill=W.CALC)
        txt(ws, r, 4, t["leased"], fmt=NUM, fill=W.CALC)
        txt(ws, r, 5, t["vacant"], fmt=NUM, fill=W.CALC)
        txt(ws, r, 6, f"=D{r}/C{r}", fmt=PCT,
            fill=PatternFill("solid", fgColor=B.T_BAD if low else B.T_NEUTRAL),
            font=Font(name=B.FONT, size=10, bold=low,
                      color=B.D_BAD if low else B.INK))
        for c_ in (7, 8, 9): txt(ws, r, c_, "", box=False)
        txt(ws, r, 10, t["rival"], align=W.WRAP,
            font=Font(name=B.FONT, size=9, bold=low,
                      color=B.D_BAD if low else B.INK2))
        r += 1
    p1 = r - 1
    txt(ws, r, 2, "الإجمالي", align=W.RGT, font=W.BOLD, fill=W.FS)
    for c_ in (3, 4, 5):
        txt(ws, r, c_, f"=SUM({get_column_letter(c_)}{p0}:{get_column_letter(c_)}{p1})",
            fmt=NUM, fill=W.FS, font=W.BOLD)
    txt(ws, r, 6, f"=D{r}/C{r}", fmt=PCT, fill=W.FS, font=W.BOLD)
    for c_ in (7, 8, 9): txt(ws, r, c_, "", box=False)
    txt(ws, r, 10, "", fill=W.FS)
    r += 1
    r = note(ws, r, NC,
             f"المحل هو الصيغة الوحيدة التي يملك مستأجرها مقارنةً يخسرها موقعُنا — "
             f"{shop['n']} محلاً بإشغال {shop['occ']*100:.1f}٪، بينما الصيغ التي "
             f"لا بديل لها بحركة مضمونة (أسواق · أكشاك · سيّاقة · مغاسل) بين ٤٤٪ "
             f"و٥١٪. والإشغال العام {PR['occ']*100:.0f}٪. "
             "وهذه أفضل قراءة تحتملها البيانات، ويُثبتها نهائياً جدولُ إيجارات "
             "ومبيعاتُ مستأجرين لا نملكهما بعد.")
    return ws


# ═══════════════════════════════════════════════ ⑩ الوحدات والاكتتاب
def leasing(wb):
    NC = 10
    U = X["underwrite"]
    ws = sheet(wb, "الوحدات والاكتتاب", [3, 28, 12, 10, 10, 13, 13, 4, 13, 30],
               "سعِّر من دفتر المستأجر لا من متر المربّع",
               "الوحدة ليست مساحة — هي وصولٌ إلى تيّار حركة، وقيمتها ما يستطيع "
               f"المستأجر تحويله منه · المدى {U['spread']:.0f}× بين أعلى وأدنى", NC)
    r = 4
    r = W.band(ws, r, NC, "① المُدخَل الأول — حركةٌ لكل وحدة شاغرة")
    W.header(ws, r, ["", "المحطة", "الفئة", "وحدات", "شاغرة", "زيارة/يوم",
                     "زيارة لكل وحدة شاغرة", "", "الحكم", "الفعل"])
    r += 1
    u0 = r
    med = U["median"]
    for s in U["ready"] + U["thin"]:
        above = s["per"] >= med
        fill = PatternFill("solid", fgColor=B.T_GOOD if above else B.T_BAD)
        txt(ws, r, 2, s["name"], align=W.RGT, font=W.BOLD)
        txt(ws, r, 3, s["cat"], font=Font(name=B.FONT, size=9, color=B.INK2))
        txt(ws, r, 4, s["units"], fmt=NUM, fill=W.CALC)
        txt(ws, r, 5, s["vacant"], fmt=NUM, fill=W.CALC)
        txt(ws, r, 6, s["vpd"], fmt=NUM, fill=W.CALC)
        txt(ws, r, 7, f"=F{r}/E{r}", fmt=NUM, fill=fill,
            font=Font(name=B.FONT, size=10, bold=True,
                      color=B.D_GOOD if above else B.D_BAD))
        txt(ws, r, 8, "", box=False)
        txt(ws, r, 9, "فوق الوسيط" if above else "دون الوسيط", fill=fill,
            font=Font(name=B.FONT, size=9, bold=True,
                      color=B.D_GOOD if above else B.D_BAD))
        txt(ws, r, 10, "الحركة موجودة — سعِّر منها" if above else
            "لا يُصلحها خفض إيجار — مشكلة تحجيم أو ملكية مهمّة", align=W.WRAP,
            font=Font(name=B.FONT, size=8, color=B.INK2))
        r += 1
    u1 = r - 1
    txt(ws, r, 2, "الوسيط", align=W.RGT, font=W.BOLD, fill=W.FS)
    for c_ in (3, 4, 5, 6): txt(ws, r, c_, "", fill=W.FS)
    txt(ws, r, 7, f"=MEDIAN(G{u0}:G{u1})", fmt=NUM, fill=W.FS, font=W.BOLD)
    txt(ws, r, 8, "", box=False)
    for c_ in (9, 10): txt(ws, r, c_, "", fill=W.FS)
    r += 1
    ws.conditional_formatting.add(f"G{u0}:G{u1}",
                                  DataBarRule(start_type="num", start_value=0,
                                              end_type="max", color=B.ORANGE, showValue=True))
    r = note(ws, r, NC,
             f"المدى {U['spread']:.0f} ضعفاً — من {U['ready'][0]['per']:,.0f} زيارة/يوم "
             f"لكل وحدة شاغرة في {U['ready'][0]['name'][:14]} إلى "
             f"{U['thin'][-1]['per']:,.0f} في {U['thin'][-1]['name'][:14]}. "
             "فوق الوسيط: الحركة موجودة فسعِّر منها. دونه: لا يُصلحها خفض إيجار.")
    r += 1

    r = W.band(ws, r, NC, "② التغطية — كم وحدةً نملك مُدخَلها الأول أصلاً؟")
    W.header(ws, r, ["", "الفئة", "شاغرة", "مقيسة", "التغطية", "", "", "", "",
                     "لماذا وما الفعل"])
    r += 1
    c0 = r
    for c_ in U["coverage"]:
        txt(ws, r, 2, c_["cat"], align=W.RGT, font=W.BOLD)
        txt(ws, r, 3, c_["vacant"], fmt=NUM, fill=W.CALC)
        txt(ws, r, 4, c_["measured"], fmt=NUM, fill=W.CALC)
        low = c_["share"] < 0.5
        txt(ws, r, 5, f"=D{r}/C{r}", fmt=PCT,
            fill=PatternFill("solid", fgColor=B.T_BAD if low else B.T_GOOD),
            font=Font(name=B.FONT, size=10, bold=True,
                      color=B.D_BAD if low else B.D_GOOD))
        for c2 in range(6, 10): txt(ws, r, c2, "", box=False)
        txt(ws, r, 10, c_["why"], align=W.WRAP,
            font=Font(name=B.FONT, size=9, color=B.INK2))
        r += 1
    c1 = r - 1
    txt(ws, r, 2, "الإجمالي", align=W.RGT, font=W.BOLD, fill=W.FS)
    for c_ in (3, 4): txt(ws, r, c_, f"=SUM({get_column_letter(c_)}{c0}:{get_column_letter(c_)}{c1})",
                          fmt=NUM, fill=W.FS, font=W.BOLD)
    txt(ws, r, 5, f"=D{r}/C{r}", fmt=PCT, fill=W.FS, font=W.BOLD)
    for c2 in range(6, 10): txt(ws, r, c2, "", box=False)
    txt(ws, r, 10, "", fill=W.FS)
    r += 1
    r = note(ws, r, NC,
             f"التغطية {U['measured_share']*100:.0f}٪ فقط: {U['measured']} من "
             f"{U['total_vacant']} وحدة شاغرة نملك مُدخَلها الأول. "
             "وحين تجهل قيمة الحركة لا تراهن عليها — شارِك فيها: نموذج «نسبة من "
             "المبيعات» يُسعّر الموقع من دفتر المستأجر تلقائياً وينقل خطر التقدير "
             "الخاطئ عنّا.",
             fill=PatternFill("solid", fgColor=B.T_GOLD),
             font=Font(name=B.FONT, size=9, bold=True, color=B.D_GOLD), h=36)
    r += 1

    r = W.band(ws, r, NC, "③ ورقة الاكتتاب — أربعة مُدخَلات قبل النطق برقم")
    W.header(ws, r, ["", "المُدخَل", "عندنا؟", "", "", "", "", "", "", "المصدر أو المالك"])
    r += 1
    owner = {2: "التأجير — مسح تحويل لكل صيغة",
             3: "التأجير — إقرار مبيعات المستأجرين شهرياً"}
    for i, it in enumerate(U["inputs"]):
        name, detail, ok = it[0], it[1], bool(it[2])
        txt(ws, r, 2, name, align=W.RGT, font=W.BOLD)
        merge(ws, r, 3, 9, ("✓ عندنا — " if ok else "✕ ليس عندنا — ") + detail,
              align=W.RGT,
              fill=PatternFill("solid", fgColor=B.T_GOOD if ok else B.T_BAD),
              font=Font(name=B.FONT, size=10, bold=True,
                        color=B.D_GOOD if ok else B.D_BAD))
        txt(ws, r, 10, "كاش إن" if ok else owner.get(i, "التأجير"), align=W.WRAP,
            font=Font(name=B.FONT, size=9, color=B.INK2))
        r += 1
    r = note(ws, r, NC,
             "المُدخَلان الأولان عندنا، والأخيران لا. ولا يُملأ الفراغ بتقدير يبدو "
             "كقياس — بل يُكتب ما ينقص ومن يملك سدّه.")
    return ws


# ═══════════════════════════════════════════════ ⑪ المنهجية والثغرات
def method(wb):
    NC = 8
    ws = sheet(wb, "المنهجية والثغرات", [3, 26, 20, 20, 4, 26, 20, 26],
               "من أين جاء كل رقم — وما الذي لا نعرفه",
               "ولا يُدَّعى مسحٌ لم يُجرَ: حيث تنقص المعرفة يُكتب ما ينقص بالضبط "
               "ومن يملك سدّه", NC)
    r = 4
    r = W.band(ws, r, NC, "① المصادر")
    W.header(ws, r, ["", "المصدر", "النطاق", "ما يعطيه", "", "التحفّظ", "", ""])
    r += 1
    src = [
        ("كاش إن — الشبكة", f"{NET['stations']} محطة · {NET['days']:,} يوم-محطة",
         "زيارات ولترات ومزيج وقود وقنوات دفع وشكل ساعات",
         f"{NET['raw_rows']:,} صفّاً · محطات فُتحت أو أُغلقت داخل المدى فلا تُقارَن "
         "مجاميعها"),
        ("كاش إن — العمرة اليومي", f"{D['days']} يوماً كاملاً",
         "الساعة وشريحة الفاتورة وزمن الخدمة — المحطة الوحيدة",
         f"يوم {D['partial']} ناقص فاستُبعد · {D['dupes']:,} مكرَّراً و"
         f"{D['skipped']:,} صفّاً متجاوَزاً من {D['rows']:,}"),
        ("ملف ساعات الشبكة", f"{H['ref_n']} محطة حيوية",
         "شكل الساعات المرجعي الذي تُقاس عليه الفجوة",
         "وسيطٌ لا مستهدف — ونصف المحطات فوقه بالتعريف"),
        ("قائمة الدخل يناير–يوليو ٢٠٢٦",
         f"{LC['pl']['total_stations']} محطة · {LC['pl']['total_litres']/1e6:,.0f} مليون لتر",
         "كلفة اللتر بأربع طبقات وبتفصيل نموذج العمل",
         "شبكيّة لا محطّية — ولا نملك تفصيل كل محطة"),
        ("تقرير ربحية الربع الثاني", f"{len(LC['op']['rows'])} محطة مشغّلة",
         "وحده يفصّل داخل «مصاريف التشغيل»: أجور · مرافق",
         "⛔ الأجر فيه ٣٬٥٠٠ ريالاً ثابتة لكل عامل — سعرٌ لا قياس"),
        ("انحدار هامش ١٩ محطة", "من تقرير المالية",
         "هامش البنزين ١٢٫٦٦ هللة والديزل ٤٫٤٧",
         "⛔ استقراء: أقصى حصة ديزل مرصودة ٤٣٪ — وليسا جدولاً معتمداً"),
        ("ملفات الطاقم والعمليات", f"{SG['total_staff']} عاملاً · {SG['n']} محطة",
         "توزيع العمّال على الورديتين وحِمل كلٍّ منهما",
         f"⛔ {SG['no_worker']} محطة بلا بيان عامل · والعمرة ٢٠ و٢٦ و٣٢ في ثلاثة ملفات"),
        ("مسح المنافسين الميداني", f"{X['retail']['n_sample']} موقعاً حول محطاتنا",
         "العلامة والمسافة والتقييم — لوقود الأفراد وحده",
         "لا يوجد مسح مقابل للعقار ولا للإعلان ولا للإكسسوارات"),
        ("سجلّ الوحدات التجارية", f"{P['units']['n']} محطة · {P['units']['total']} وحدة",
         "الفئة والمؤجَّر والشاغر", "لا جدول إيجارات ولا مبيعات مستأجرين"),
        ("سجلّ الشكاوى", f"{D['complaints']['net']} شكوى",
         f"{D['complaints']['n']} منها للعمرة ({D['complaints']['share']*100:.0f}٪)",
         "الشكاوى ترتفع بالحجم — والحصة هنا تفوق حصة الحجم"),
    ]
    for a, b, c_, d_ in src:
        txt(ws, r, 2, a, align=W.RGT, font=W.BOLD)
        txt(ws, r, 3, b, align=W.WRAP, font=Font(name=B.FONT, size=9, color=B.INK2))
        merge(ws, r, 4, 5, c_, align=W.RGT,
              font=Font(name=B.FONT, size=9, color=B.INK2))
        gate = d_.startswith("⛔")
        merge(ws, r, 6, 8, d_, align=W.RGT,
              fill=PatternFill("solid", fgColor=B.T_GOLD) if gate else None,
              font=Font(name=B.FONT, size=9, bold=gate,
                        color=B.D_GOLD if gate else B.INK2))
        ws.row_dimensions[r].height = 32
        r += 1
    r += 1

    r = W.band(ws, r, NC, "② التعاريف — حتى لا يُقاس رقمان بتعريفين")
    W.header(ws, r, ["", "المصطلح", "التعريف", "", "", "لماذا هكذا", "", ""])
    r += 1
    defs = [
        ("الفجوة", "ما نحن دونه من شكل تصنيفنا في ساعات ذروتنا الثماني وحدها",
         "النقص في ساعات الهدوء اختلاف طلبٍ لا فاقد — فلا يُطالَب به"),
        ("الوردية", "صباحية ٠٠:٠٠–١١:٥٩ · مسائية ١٢:٠٠–٢٣:٥٩",
         "حدود التقرير نفسه — لا حدودٌ من عندنا"),
        ("هامش المساهمة", "الإيراد ناقص شراء الوقود — قبل أي مصروف تشغيل",
         "هو وحده ما يتحرّك باللتر الإضافي"),
        ("خط الأساس", "الأسابيع الأربعة السابقة في المحطة نفسها",
         "القياس على ذروةٍ موسمية أو على عامٍ مضى يخلط الموسم بالأثر"),
        ("البوابة", "حالةٌ تُبطل نسبة أي نتيجة إلى حملة",
         "محطةٌ عليها بوابة لا تدخل حملة قبل إقفالها"),
        ("الربع الثالث للعمرة", "١ يوليو – ٢٠ سبتمبر · ٨٢ يوماً",
         "ناقصٌ فتُقارَن المعدّلات اليومية لا المجاميع"),
    ]
    for a, b, c_ in defs:
        txt(ws, r, 2, a, align=W.RGT, font=W.BOLD)
        merge(ws, r, 3, 5, b, align=W.RGT)
        merge(ws, r, 6, 8, c_, align=W.RGT,
              font=Font(name=B.FONT, size=9, color=B.INK2))
        ws.row_dimensions[r].height = 26
        r += 1
    r += 1

    r = W.band(ws, r, NC, "③ الثغرات — ما ينقص بالضبط ومن يملك سدّه")
    W.header(ws, r, ["", "الناقص", "المالك", "الحال اليوم", "", "أثره على الرقم", "", ""])
    r += 1
    gaps = [
        ("جدول ورديات معتمد", "العمليات",
         "ثلاثة ملفات تعطي أعداداً مختلفة للمحطة الواحدة",
         "ملزم قبل أي مستهدف وردية — والمستهدف يُحمَّل على وردية"),
        ("جدول هامش المنتجات", "المالية",
         "١٢٫٦٦ و٤٫٤٧ هللة استُنتجا بانحدار لا بجدول",
         "يغيّر قيمة كل فجوة في الملف — والديزل ٣٣٪ في محطاتٍ كبرى"),
        ("رسوم البطاقات", "المالية", "لا نعرف الرسم ولا من يتحمّله",
         "يُقتطع من هامش كل لتر إضافي ولا يظهر في أي حساب هنا"),
        ("جرد الخزانات", "العمليات", "لا مطابقة بين المُورَّد والمُباع",
         "يفصل فجوة المطابقة عن فجوة البيع"),
        ("نموذج عمل كل محطة", "المالية + العمليات",
         "العمرة استثمارية وإيجارية وتشغيلية في ثلاثة ملفات",
         "يغيّر أساس الكلفة كلّه"),
        ("جدول إيجارات ومبيعات المستأجرين", "التأجير",
         "لا نملك دفتر المستأجر الذي نُسعّر منه",
         "المُدخَلان الثالث والرابع في ورقة الاكتتاب"),
        ("مسح منافسين للعقار والإعلان", "التجاري",
         "المسح الميداني لوقود الأفراد وحده",
         "لا حصص سوق لهذين المنتجين — ولا تُدَّعى"),
        ("أجرة النقل والتوريد", "سلسلة الإمداد", "غير محسوم من يتحمّلها",
         "بندٌ كامل خارج كلفة اللتر عندنا"),
    ]
    for a, owner, now, eff in gaps:
        txt(ws, r, 2, a, align=W.RGT, font=W.BOLD)
        txt(ws, r, 3, owner, fill=PatternFill("solid", fgColor=B.T_GOLD),
            font=Font(name=B.FONT, size=9, bold=True, color=B.D_GOLD))
        merge(ws, r, 4, 5, now, align=W.RGT,
              font=Font(name=B.FONT, size=9, color=B.INK2))
        merge(ws, r, 6, 8, eff, align=W.RGT,
              font=Font(name=B.FONT, size=9, color=B.INK2))
        ws.row_dimensions[r].height = 26
        r += 1
    r = note(ws, r, NC,
             "وأهمّها الأول: جدول ورديات معتمد. فالمستهدف يُحمَّل على وردية، ولا "
             "يُحمَّل على وردية عدد عمّالها ثلاثة أرقام في ثلاثة ملفات.",
             fill=PatternFill("solid", fgColor=B.T_GOLD),
             font=Font(name=B.FONT, size=10, bold=True, color=B.D_GOLD))
    return ws


def build():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    summary(wb)
    stations(wb)          # يجب أن تسبق الحاسبة — منها يُبنى مرجع البحث
    hourshapes(wb)        # وكذلك شكل الساعات — منه تُشتقّ الفجوة أمام القارئ
    calculator(wb)
    quarters_net(wb)
    mk007(wb)
    hours(wb)
    costsheet(wb)
    plan(wb)
    competitors(wb)
    leasing(wb)
    method(wb)
    # الحاسبة ثانيةً في الترتيب ليجدها القارئ أولاً بعد الملخص
    wb.move_sheet(S_CALC, offset=-1)
    wb.properties.title = "الحاسبة التجارية — درب"
    wb.properties.creator = "القسم التجاري — درب"
    wb.save(OUT)
    return OUT


if __name__ == "__main__":
    p = build()
    print("كُتب", p)
