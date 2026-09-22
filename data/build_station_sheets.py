# -*- coding: utf-8 -*-
"""شيتٌ لكل محطة — الأداء وحاسبة اللتر والمستهدف الشهري

   في كل شيت ثلاثة أقسام لا رابع لها:
   ① الأداء     — الربع الأول والثاني من تقرير الربحية، والإيراد والصافي
                  وهامش الوقود من تقرير المالية، وحركتها من كاش إن
   ② حاسبة اللتر — قيمة اللتر · هامش الربح · المصاريف · الصافي · بصيغٍ حيّة
   ③ المستهدف    — اثنا عشر شهراً، مبنيّةً على أداء المحطة نفسها لا على رغبة

   وما لا مصدر له يُكتب شرطةً ويُقال من يملك سدّه — لا يُملأ بتقدير.

   يُشغَّل من جذر المستودع:  PYTHONPATH=data python3 data/build_station_sheets.py
"""
import calendar, csv, json
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, DataBarRule

import brand as B
import worker_model as W
import sales_plan as S

PLAN = "data/sales-plan.json"
Q2 = "outlets/data/q2-stations.csv"
Q2C = "outlets/data/q2-costs.csv"
CFO = "outlets/data/cfo-stations.csv"
NETS = "outlets/data/network-sales.csv"
SHIFTS = "outlets/data/worker-shifts.csv"
OUT = "docs/شيت-كل-محطة.xlsx"

MONEY = '#,##0;[Red](#,##0)'
MONEY2 = '#,##0.00;[Red](#,##0.00)'
NUM = '#,##0'
NUM2 = '#,##0.00'
PCT = '0.0%'
PCT2 = '0.00%'
HAL = '0.00" هللة"'

# اثنا عشر شهراً تبدأ من أكتوبر ٢٠٢٦ — بأيامها الفعلية لا بثلاثين
MONTHS = [("أكتوبر ٢٠٢٦", 2026, 10), ("نوفمبر ٢٠٢٦", 2026, 11),
          ("ديسمبر ٢٠٢٦", 2026, 12), ("يناير ٢٠٢٧", 2027, 1),
          ("فبراير ٢٠٢٧", 2027, 2), ("مارس ٢٠٢٧", 2027, 3),
          ("أبريل ٢٠٢٧", 2027, 4), ("مايو ٢٠٢٧", 2027, 5),
          ("يونيو ٢٠٢٧", 2027, 6), ("يوليو ٢٠٢٧", 2027, 7),
          ("أغسطس ٢٠٢٧", 2027, 8), ("سبتمبر ٢٠٢٧", 2027, 9)]
RAMP = 6          # أشهر التدرّج حتى بلوغ المستهدف كاملاً

# مصاريف تشغيل المحطة بنموذج عملها — من قائمة الدخل يناير–يوليو ٢٠٢٦
#   ولا يُحمَّل على محطة امتياز ما يُحمَّل على محطة تشغيلية: الفرق ٢٤٦ ضعفاً
MODEL_CPL = {"استثماري": 11.08, "إيجاري": 5.63, "تشغيلي": 9.84, "امتياز": 0.04}


def bases_for(s):
    """أسُس المصاريف لهذه المحطة — أولُها نموذج عملها إن عُرف"""
    out = []
    m = s["model"]
    if m in MODEL_CPL:
        out.append((f"مصاريف نموذج «{m}» — {MODEL_CPL[m]:.2f} هللة/لتر",
                    MODEL_CPL[m],
                    f"مصاريف تشغيل محطات {m} في قائمة الدخل ÷ لتراتها — "
                    "وهو الأقرب لهذه المحطة"))
    out.append(("متوسط محطاتنا — ٩٫١٥ هللة/لتر", S.OPEX_NET * 100,
                "مصاريف تشغيل محطاتنا الـ٤٩ (دون الامتياز) ÷ لتراتها"))
    out.append(("المحمَّل الكامل — ١٣٫٥٦ هللة/لتر", (S.OPEX_NET + S.LOADED) * 100,
                "يُضاف إليها بيعٌ وتسويق وإدارة وفوائد — لقرار رأس المال "
                "لا للتشغيل"))
    out.append(("المقيس في تقرير الربع الثاني", s["q2_cpl"],
                "⛔ أجورُه ٣٬٥٠٠ ريالاً ثابتة لكل عامل — أجرٌ شهري طُبِّق مرة "
                "بدل ثلاث، فمصاريفه نصف ما في القائمة · يصلح للمقارنة بين "
                "المحطات لا لمستوى الكلفة"
                if s["q2_cpl"] else
                "⛔ لا أجور ولا مرافق مسجَّلة لهذه المحطة — لا تختر هذا الأساس"))
    return out


def _f(v, d=0.0):
    try: return float(str(v).replace(",", "").strip())
    except (TypeError, ValueError): return d


def load():
    P = json.load(open(PLAN, encoding="utf-8"))
    plan = {s["code"]: s for s in P["stations"]}
    contr = {c["code"]: c for c in P["contracts"]["stations"]}

    def idx(path, key, enc="utf-8"):
        return {r[key]: r for r in csv.DictReader(open(path, encoding=enc))
                if r.get(key)}

    q2 = {c: r for c, r in idx(Q2, "code").items() if _f(r["q2_vol"])}
    cost = idx(Q2C, "code")
    cfo = idx(CFO, "code")
    net = idx(NETS, "code")
    shift = idx(SHIFTS, "الرمز")

    out = []
    for code in sorted(set(q2) | set(cfo)):
        p, q, c, k = plan.get(code), q2.get(code), cfo.get(code), cost.get(code)
        n, sh = net.get(code, {}), shift.get(code, {})
        name = (p or {}).get("name") or (q or {}).get("name") \
            or (c or {}).get("name") or code
        s = dict(code=code, name=str(name).strip(),
                 region=(p or {}).get("region", "—"),
                 seg=(p or {}).get("seg", "—"),
                 model=(c or {}).get("model", "—"),
                 kind=(c or {}).get("kind", "—"))
        # ① الأداء
        s["q2"] = dict(
            sales1=_f(q and q["q1_sales"]), sales2=_f(q and q["q2_sales"]),
            vol1=_f(q and q["q1_vol"]), vol2=_f(q and q["q2_vol"]),
            net1=_f(q and q["q1_net"]), net2=_f(q and q["q2_net"]),
            mar1=_f(q and q["q1_margin"]), mar2=_f(q and q["q2_margin"]),
            staff1=_f(q and q["q1_staff"]), staff2=_f(q and q["q2_staff"]),
            terms=(q or {}).get("terms", ""),
            wages=_f(k and k["q2_wages"]), util=_f(k and k["q2_util"]),
        ) if q else None
        ct = contr.get(code)
        s["contract"] = dict(act=ct["act"], std=ct["std"], gap=ct["gap"],
                             kind=ct["kind"]) if ct else None
        s["cfo"] = dict(rev_k=_f(c and c["rev_k"]), net_k=_f(c and c["net_k"]),
                        net_pct=_f(c and c["net_pct"]),
                        fuel_margin=_f(c and c["fuel_margin"])) if c else None
        s["plan"] = dict(vpd=p["vpd"], lpd=p["lpd"], lpv=p["lpv"],
                         diesel=p["diesel"], inv=p["inv"], peak=p["peak"],
                         gap_peak=p["gap_peak"], sar_txn=p["sar_txn"],
                         conds=p["conds"]) if p else None
        s["staff"] = int(_f(sh.get("عدد العمّال")))
        s["days_net"] = int(_f(n.get("days")))
        # خط الأساس والسعر — ومعهما أساس الضريبة صراحةً:
        #   كاش إن يسجّل المبيعات شاملةً الضريبة، وتقارير المالية قبلها.
        #   وهامش الوقود نسبةٌ من الإيراد قبل الضريبة، فلا يُضرب في سعرٍ شامل.
        if s["q2"] and s["q2"]["vol2"]:
            s["base_lpd"] = s["q2"]["vol2"] / 91.0
            s["base_src"] = "حجم الربع الثاني ÷ ٩١ يوماً"
        elif s["plan"]:
            s["base_lpd"] = s["plan"]["lpd"]
            s["base_src"] = "معدّل كاش إن للنافذة المقيسة"
        else:
            s["base_lpd"] = 0.0; s["base_src"] = "—"
        if s["q2"] and s["q2"]["vol2"] and s["q2"]["sales2"]:
            s["price_ex"] = s["q2"]["sales2"] / s["q2"]["vol2"]
            s["price_src"] = ("مبيعات الربع الثاني ÷ حجمه (وهي قبل الضريبة) "
                              "× ١٫١٥ لتصير سعر مضخة")
        elif s["plan"]:
            s["price_ex"] = p["rev"] / (p["vol"] or 1) / S.VAT
            s["price_src"] = ("إيراد كاش إن ÷ لتراته ÷ ١٫١٥ — وكاش إن يسجّل "
                              "شاملاً الضريبة" if not (s["q2"] and s["q2"]["vol2"])
                              else "⛔ لا مبيعات مسجَّلة لها في تقرير الربع الثاني "
                              "— فالسعر من كاش إن")
        else:
            s["price_ex"] = 0.0; s["price_src"] = "⛔ لا سعر مقيس لهذه المحطة"
        s["price"] = s["price_ex"] * S.VAT          # سعر المضخة شاملاً
        # هامش الوقود: المقيس من تقرير المالية أولاً، وإلا مزيج المحطة
        if s["cfo"] and s["cfo"]["fuel_margin"]:
            s["fm"] = s["cfo"]["fuel_margin"]
            s["fm_src"] = "هامش الوقود المقيس في تقرير المالية"
        elif s["plan"]:
            d = s["plan"]["diesel"]
            s["fm"] = (S.MARGIN_PETROL * (1 - d) + S.MARGIN_DIESEL * d) \
                / (s["price_ex"] or 1)
            s["fm_src"] = ("مستنتج من مزيج المحطة — بنزين ١٢٫٦٦ هللة "
                           "وديزل ٤٫٤٧ · ⛔ انحدار لا جدول معتمد")
        else:
            s["fm"] = S.MARGIN / (s["price_ex"] or 1) if s["price_ex"] else 0.0
            s["fm_src"] = "⛔ متوسط الشبكة — لا قياس لهذه المحطة"
        # الإضافة المستهدفة
        if s["plan"] and s["plan"]["gap_peak"] > 0:
            s["add_lpd"] = s["plan"]["gap_peak"] * S.CLOSE * s["plan"]["lpv"]
            s["add_src"] = (f"فجوة ساعات الذروة {s['plan']['gap_peak']:.1f} "
                            f"معاملة/يوم × {S.CLOSE*100:.0f}٪ إغلاقاً × "
                            f"{s['plan']['lpv']:.1f} لتراً للزيارة")
        else:
            s["add_lpd"] = 0.0
            s["add_src"] = "لا فجوة ذروة مقيسة لهذه المحطة — المستهدف ثباتُ الأداء"
        # مصاريف تقرير الربع الثاني لكل لتر (أجور + مرافق) — إن وُجدت
        if s["q2"] and s["q2"]["vol2"] and (s["q2"]["wages"] or s["q2"]["util"]):
            s["q2_cpl"] = (s["q2"]["wages"] + s["q2"]["util"]) / s["q2"]["vol2"] * 100
        else:
            s["q2_cpl"] = 0.0
        s["bases"] = bases_for(s)
        s["cpl_def"] = s["bases"][0][1]
        out.append(s)
    return P, out


# ═══════════════════════════════════════════════ أدوات
NC = 9


def txt(ws, r, c, v, *, fmt=None, font=None, fill=None, align=None, box=True):
    x = ws.cell(r, c, v)
    x.font = font or W.BLACK
    x.alignment = align or W.CTR
    if fill: x.fill = fill
    if fmt: x.number_format = fmt
    if box: x.border = W.BOX
    return x


def merge(ws, r, c1, c2, v, *, font=None, fill=None, align=None, fmt=None):
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    x = txt(ws, r, c1, v, font=font, fill=fill, align=align, fmt=fmt)
    for c in range(c1 + 1, c2 + 1):
        cc = ws.cell(r, c); cc.border = W.BOX
        if fill: cc.fill = fill
    return x


def note(ws, r, text, *, tone=None, h=26, nc=NC):
    fl = PatternFill("solid", fgColor={"gold": B.T_GOLD, "band": B.T_ORANGE,
                                       "bad": B.T_BAD}.get(tone, B.BG))
    col = {"gold": B.D_GOLD, "band": B.ORANGE, "bad": B.D_BAD}.get(tone, B.INK2)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=nc)
    c = ws.cell(r, 2, text)
    c.font = Font(name=B.FONT, size=9, bold=tone is not None, color=col)
    c.alignment = W.WRAP; c.fill = fl; c.border = W.BOX
    for j in range(3, nc + 1):
        cc = ws.cell(r, j); cc.border = W.BOX; cc.fill = fl
    ws.row_dimensions[r].height = h
    return r + 1


def band(ws, r, text, nc=NC):
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=nc)
    c = ws.cell(r, 2, text)
    c.font = Font(name=B.FONT, size=12, bold=True, color=B.ORANGE)
    c.fill = PatternFill("solid", fgColor=B.T_ORANGE)
    c.alignment = W.RGT; c.border = W.BOX
    for j in range(3, nc + 1):
        cc = ws.cell(r, j); cc.border = W.BOX
        cc.fill = PatternFill("solid", fgColor=B.T_ORANGE)
    ws.row_dimensions[r].height = 22
    return r + 1


def sheet_name(s, used):
    nm = f"{s['code']} {s['name']}"
    for ch in "[]:*?/\\": nm = nm.replace(ch, " ")
    nm = nm[:31].strip()
    base, i = nm, 2
    while nm in used:
        nm = f"{base[:28]}~{i}"; i += 1
    used.add(nm)
    return nm


def _mh(ws, r, c1, c2):
    """دمج خلايا الترويسة — ليجلس العنوان حيث يبدأ نصُّه"""
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)


# ═══════════════════════════════════════════════ شيت المحطة
def station(wb, s):
    ws = wb.create_sheet(s["sheet"])
    W.setup(ws, [3, 30, 15, 15, 15, 15, 15, 15, 34], freeze="A5")
    bits = [x for x in (s["region"], s["seg"]) if x and x != "—"]
    if s["model"] != "—": bits.append(f"نموذج العمل {s['model']}")
    if s["kind"] != "—": bits.append(f"{s['kind']} في تقرير المالية")
    bits.append("مصدر الأداء: " + ("تقرير ربحية الربع الثاني" if s["q2"]
                                   else "تقرير المالية وكاش إن"))
    W.title(ws, f"{s['name']}  ·  {s['code']}", " · ".join(bits), NC)
    W.rule(ws, 3, NC)
    r = 4

    # ── ① الأداء
    r = band(ws, r, "① الأداء — كما وصلنا من التقارير")
    if s["q2"]:
        q = s["q2"]
        W.header(ws, r, ["", "البند", "الربع الأول", "الربع الثاني", "التغيّر",
                         "المصدر", "", "", ""])
        _mh(ws, r, 6, 9)
        r += 1
        p0 = r
        prows = [
            ("المبيعات", q["sales1"], q["sales2"], MONEY, "ريال"),
            ("الحجم", q["vol1"], q["vol2"], NUM, "لتر"),
            ("صافي المحطة", q["net1"], q["net2"], MONEY, "ريال"),
            ("هامش الصافي", q["mar1"] / 100, q["mar2"] / 100, PCT, "٪"),
            ("عدد العمّال", q["staff1"], q["staff2"], NUM, "عامل"),
        ]
        for lab, a, b, fm, u in prows:
            txt(ws, r, 2, lab, align=W.RGT, font=W.BOLD)
            txt(ws, r, 3, a if a else "—", fmt=fm if a else None, fill=W.CALC)
            txt(ws, r, 4, b if b else "—", fmt=fm if b else None, fill=W.CALC)
            if a and b:
                txt(ws, r, 5, f"=D{r}/C{r}-1" if fm != PCT else f"=D{r}-C{r}",
                    fmt=PCT if fm != PCT else '+0.0%;-0.0%', fill=W.CALC,
                    font=W.BOLD)
            else:
                txt(ws, r, 5, "—")
            merge(ws, r, 6, 9, "تقرير ربحية الربع الثاني", align=W.RGT,
                  font=Font(name=B.FONT, size=8, color=B.INK3))
            r += 1
        for c_ in ("E",):
            ws.conditional_formatting.add(f"{c_}{p0}:{c_}{r-1}", CellIsRule(
                operator="lessThan", formula=["0"],
                font=Font(name=B.FONT, size=10, bold=True, color=B.BAD)))
            ws.conditional_formatting.add(f"{c_}{p0}:{c_}{r-1}", CellIsRule(
                operator="greaterThan", formula=["0"],
                font=Font(name=B.FONT, size=10, bold=True, color=B.D_GOOD)))
        # الأجور والمرافق
        if q["wages"] or q["util"]:
            txt(ws, r, 2, "أجور الربع الثاني", align=W.RGT, font=W.BOLD)
            txt(ws, r, 3, "—"); txt(ws, r, 4, q["wages"], fmt=MONEY, fill=W.CALC)
            txt(ws, r, 5, "—")
            merge(ws, r, 6, 9, "⛔ ٣٬٥٠٠ ريالاً لكل عامل — سعرٌ ثابت لا قياس",
                  align=W.RGT, fill=PatternFill("solid", fgColor=B.T_GOLD),
                  font=Font(name=B.FONT, size=8, bold=True, color=B.D_GOLD))
            r += 1
            txt(ws, r, 2, "مرافق الربع الثاني", align=W.RGT, font=W.BOLD)
            txt(ws, r, 3, "—"); txt(ws, r, 4, q["util"], fmt=MONEY, fill=W.CALC)
            txt(ws, r, 5, "—")
            merge(ws, r, 6, 9, "تقرير ربحية الربع الثاني", align=W.RGT,
                  font=Font(name=B.FONT, size=8, color=B.INK3))
            r += 1
        if s["contract"]:
            ct = s["contract"]
            txt(ws, r, 2, "حصّتنا مقابل المعياري", align=W.RGT, font=W.BOLD)
            txt(ws, r, 3, ct["act"], fmt=MONEY, fill=W.CALC)
            txt(ws, r, 4, ct["std"], fmt=MONEY, fill=W.CALC)
            txt(ws, r, 5, f"=C{r}-D{r}", fmt=MONEY, font=W.BOLD,
                fill=PatternFill("solid",
                                 fgColor=B.T_BAD if ct["gap"] < 0 else B.T_GOOD))
            merge(ws, r, 6, 9, ct["kind"], align=W.RGT,
                  font=Font(name=B.FONT, size=8, color=B.INK3))
            r += 1
            r = note(ws, r, f"شروط العقد: {q['terms']}")
    else:
        r = note(ws, r, "لا تقرير ربعيّ لهذه المحطة — تقرير الربحية يغطّي "
                 "المحطات المشغّلة وحدها. والمتاح عنها أدناه.", tone="gold")

    if s["cfo"]:
        c = s["cfo"]
        W.header(ws, r, ["", "من تقرير المالية", "القيمة", "الوحدة",
                         "الملاحظة", "", "", "", ""])
        _mh(ws, r, 5, 9)
        r += 1
        crows = [("الإيراد", c["rev_k"], NUM, "ألف ريال", "تراكمي يوليو ٢٠٢٦"),
                 ("صافي الدخل", c["net_k"], MONEY, "ألف ريال",
                  "بعد المحمَّلات والفوائد"),
                 ("هامش الصافي", c["net_pct"], PCT2, "٪", ""),
                 ("هامش الوقود", c["fuel_margin"], PCT2, "٪ من الإيراد",
                  "وهو أساس حاسبة اللتر أدناه")]
        for lab, v, fm, u, nt in crows:
            txt(ws, r, 2, lab, align=W.RGT, font=W.BOLD)
            tone = (B.T_BAD if lab == "صافي الدخل" and v < 0 else
                    B.T_GOOD if lab == "صافي الدخل" else B.T_NEUTRAL)
            txt(ws, r, 3, v, fmt=fm, fill=PatternFill("solid", fgColor=tone),
                font=Font(name=B.FONT, size=11, bold=True,
                          color=B.D_BAD if tone == B.T_BAD else
                          B.D_GOOD if tone == B.T_GOOD else B.INK))
            txt(ws, r, 4, u, font=Font(name=B.FONT, size=9, color=B.INK3))
            merge(ws, r, 5, 9, nt, align=W.RGT,
                  font=Font(name=B.FONT, size=8, color=B.INK3))
            r += 1
    if s["plan"]:
        p = s["plan"]
        W.header(ws, r, ["", "من كاش إن", "القيمة", "الوحدة", "الملاحظة",
                         "", "", "", ""])
        _mh(ws, r, 5, 9)
        r += 1
        for lab, v, fm, u in (("زيارة في اليوم", p["vpd"], NUM, "زيارة"),
                              ("لتر في اليوم", p["lpd"], NUM, "لتر"),
                              ("لتر لكل زيارة", p["lpv"], NUM2, "لتر"),
                              ("حصة الديزل", p["diesel"], PCT, "من اللترات"),
                              ("ساعة الذروة", p["peak"], NUM, "من ٢٤")):
            txt(ws, r, 2, lab, align=W.RGT, font=W.BOLD)
            txt(ws, r, 3, v, fmt=fm, fill=W.CALC)
            txt(ws, r, 4, u, font=Font(name=B.FONT, size=9, color=B.INK3))
            merge(ws, r, 5, 9, "")
            r += 1
    r += 1

    # ── ② حاسبة اللتر
    r = band(ws, r, "② حاسبة اللتر — قيمته وهامشه ومصاريفه وصافيه")
    W.header(ws, r, ["", "البند", "القيمة", "الوحدة", "من أين جاءت",
                     "", "", "", ""])
    _mh(ws, r, 5, 9)
    r += 1
    L0 = r
    # جدول أسُس المصاريف — يُكتب أولاً ليُشار إليه
    lut = L0 + 21
    rowsL = [
        ("لترات الشهر", round(s["base_lpd"] * 30.4), NUM, "لتر/شهر", "inp",
         f"افتراضيّها معدّل الأداء: {s['base_src']} × ٣٠٫٤ يوماً"),
        ("سعر المضخة — شاملاً الضريبة", s["price"], MONEY2, "ريال/لتر", "inp",
         s["price_src"]),
        ("هامش الوقود", s["fm"], PCT2, "٪ من الإيراد قبل الضريبة", "inp",
         s["fm_src"]),
        ("① قيمة اللتر — قبل الضريبة", None, MONEY2, "ريال/لتر", "key",
         "الصيغة: سعر المضخة ÷ ١٫١٥ · وهي قيمته التي يُحسب عليها الهامش"),
        ("② هامش الربح للتر", None, HAL, "هللة/لتر", "key",
         "الصيغة: قيمة اللتر قبل الضريبة × هامش الوقود · هامش المساهمة "
         "قبل أي مصروف"),
        ("أساس المصاريف", s["bases"][0][0], None, "اختر", "inp",
         "ثلاثة أُسُس ولكلٍّ مصدره — الجدول أسفل الحاسبة"),
        ("③ المصاريف للتر", None, HAL, "هللة/لتر", "key",
         "من الأساس المختار أعلاه"),
        ("④ صافي اللتر", None, HAL, "هللة/لتر", "key",
         "الصيغة: الهامش − المصاريف · وهو ما يبقى فعلاً من كل لتر"),
    ]
    for i, (lab, v, fm, u, kind, src) in enumerate(rowsL):
        rr = r + i
        key = kind == "key"
        fill = (W.FI if kind == "inp" else
                PatternFill("solid", fgColor=B.T_GOOD) if key else W.CALC)
        txt(ws, rr, 2, lab, align=W.RGT,
            font=Font(name=B.FONT, size=11 if key else 10, bold=True))
        if kind == "inp":
            W.inp(ws, rr, 3, fm or "General", v)
        else:
            f = {3: f"=$C${L0+1}/{S.VAT}",
                 4: f"=$C${L0+3}*$C${L0+2}*100",
                 6: f"=VLOOKUP($C${L0+5},$B${lut}:$C${lut+3},2,FALSE)",
                 7: f"=$C${L0+4}-$C${L0+6}"}[i]
            txt(ws, rr, 3, f, fmt=fm, fill=fill,
                font=Font(name=B.FONT, size=13 if key else 10, bold=key,
                          color=B.D_GOOD if key else B.INK))
        txt(ws, rr, 4, u, font=Font(name=B.FONT, size=9, color=B.INK3))
        gate = "⛔" in src
        merge(ws, rr, 5, 9, src, align=W.RGT,
              fill=PatternFill("solid", fgColor=B.T_GOLD) if gate else None,
              font=Font(name=B.FONT, size=8, bold=gate,
                        color=B.D_GOLD if gate else B.INK3))
    LMON, LPRICE, LFM = f"$C${L0}", f"$C${L0+1}", f"$C${L0+2}"
    LMARG, LBASIS, LCOST, LNET = (f"$C${L0+4}", f"$C${L0+5}",
                                  f"$C${L0+6}", f"$C${L0+7}")
    r += len(rowsL)
    # بالريال للشهر
    W.header(ws, r, ["", "وبالريال لشهرٍ كامل", "في الشهر", "الوحدة",
                     "الصيغة", "", "", "", ""])
    _mh(ws, r, 5, 9)
    r += 1
    M0 = r
    for lab, f, tone, why in (
            ("المبيعات قبل الضريبة", f"={LMON}*$C${L0+3}", None,
             "لترات الشهر × قيمة اللتر قبل الضريبة"),
            ("وشاملةً الضريبة", f"=C{M0}*{S.VAT}", None,
             "وهو ما يُسجَّل على المضخة"),
            ("هامش الربح", f"={LMON}*{LMARG}/100", "good",
             "لترات الشهر × هامش اللتر"),
            ("المصاريف", f"=-{LMON}*{LCOST}/100", "bad",
             "لترات الشهر × مصاريف اللتر"),
            ("الصافي", f"=C{M0+2}+C{M0+3}", "key", "الهامش − المصاريف")):
        key = tone == "key"
        fill = PatternFill("solid", fgColor={
            "good": B.T_GOOD, "bad": B.T_BAD, "key": B.T_BAND}.get(tone, B.T_NEUTRAL))
        col = {"good": B.D_GOOD, "bad": B.D_BAD, "key": B.INK}.get(tone, B.INK)
        txt(ws, r, 2, lab, align=W.RGT,
            font=Font(name=B.FONT, size=11 if key else 10, bold=True))
        txt(ws, r, 3, f, fmt=MONEY, fill=fill,
            font=Font(name=B.FONT, size=13 if key else 11, bold=True, color=col))
        txt(ws, r, 4, "ريال", font=Font(name=B.FONT, size=9, color=B.INK3))
        merge(ws, r, 5, 9, why, align=W.RGT,
              font=Font(name=B.FONT, size=8, color=B.INK3))
        r += 1
    NETM = f"$C${M0+4}"
    r += 1

    # جدول الأسُس
    while r < lut: r += 1
    W.header(ws, r - 1, ["", "أساس المصاريف", "هللة/لتر", "المصدر والتحفّظ",
                         "", "", "", "", ""])
    _mh(ws, r - 1, 4, 9)
    for i, (lab, cpl, why) in enumerate(s["bases"]):
        rr = lut + i
        txt(ws, rr, 2, lab, align=W.RGT,
            font=W.BOLD if i == 0 else W.BLACK)
        txt(ws, rr, 3, cpl, fmt=HAL,
            fill=PatternFill("solid", fgColor=B.T_GOOD) if i == 0 else W.CALC,
            font=Font(name=B.FONT, size=10, bold=i == 0,
                      color=B.D_GOOD if i == 0 else B.INK))
        gate = why.startswith("⛔")
        merge(ws, rr, 4, 9, why, align=W.RGT,
              fill=PatternFill("solid", fgColor=B.T_GOLD) if gate else None,
              font=Font(name=B.FONT, size=8, bold=gate,
                        color=B.D_GOLD if gate else B.INK3))
    dv = DataValidation(type="list",
                        formula1='"' + ",".join(b[0] for b in s["bases"]) + '"',
                        allow_blank=False)
    ws.add_data_validation(dv); dv.add(ws.cell(L0 + 5, 3))
    r = lut + len(s["bases"])
    r = note(ws, r,
             "المصاريف هنا مصاريفُ تشغيل المحطة — أجورٌ ومرافق وصيانة. "
             "أمّا شراء الوقود فمطروحٌ أصلاً داخل «هامش الوقود»، فلا يُطرح مرتين.",
             tone="gold")
    r = note(ws, r,
             "⚠ وصافي اللتر هنا صافي الوقود وحده: لا يشمل دخل الوحدات التجارية "
             "ولا خصومات الإيجار ولا محمَّلات المركز. فقد يخرج سالباً في محطةٍ "
             "رابحة يحملها دخلُ وحداتها — وصافي المحطة الكامل في القسم ① "
             "كما جاء من تقرير المالية.", tone="gold", h=32)
    r += 1

    # ── ③ المستهدف الشهري
    r = band(ws, r, "③ المستهدف الشهري — اثنا عشر شهراً")
    W.header(ws, r, ["", "الأساس", "القيمة", "الوحدة", "من أين جاءت",
                     "", "", "", ""])
    _mh(ws, r, 5, 9)
    r += 1
    T0 = r
    trows = [
        ("خط الأساس", s["base_lpd"], NUM, "لتر/يوم", "calc", s["base_src"]),
        ("الإضافة المستهدفة", s["add_lpd"], NUM, "لتر/يوم", "inp", s["add_src"]),
        ("أشهر التدرّج", RAMP, NUM, "شهر", "inp",
         "يُبلغ المستهدف كاملاً في الشهر السادس ثم يثبت"),
    ]
    for i, (lab, v, fm, u, kind, src) in enumerate(trows):
        rr = r + i
        txt(ws, rr, 2, lab, align=W.RGT, font=W.BOLD)
        if kind == "inp": W.inp(ws, rr, 3, fm, v)
        else: txt(ws, rr, 3, v, fmt=fm, fill=W.CALC)
        txt(ws, rr, 4, u, font=Font(name=B.FONT, size=9, color=B.INK3))
        merge(ws, rr, 5, 9, src, align=W.RGT,
              font=Font(name=B.FONT, size=8, color=B.INK3))
    BASE, ADD, RMP = f"$C${T0}", f"$C${T0+1}", f"$C${T0+2}"
    r += len(trows)
    W.header(ws, r, ["", "الشهر", "أيام", "لتر/يوم", "لترات الشهر",
                     "المبيعات قبل الضريبة", "هامش الربح", "المصاريف",
                     "الصافي"])
    ws.row_dimensions[r].height = 26
    r += 1
    G0 = r
    for i, (nm, yy, mm) in enumerate(MONTHS, 1):
        rr = r + i - 1
        days = calendar.monthrange(yy, mm)[1]
        txt(ws, rr, 2, nm, align=W.RGT, font=W.BOLD)
        txt(ws, rr, 3, days, fmt=NUM, fill=W.CALC)
        txt(ws, rr, 4, f"={BASE}+{ADD}*MIN(1,{i}/{RMP})", fmt=NUM, fill=W.CALC)
        txt(ws, rr, 5, f"=D{rr}*C{rr}", fmt=NUM, fill=W.CALC, font=W.BOLD)
        txt(ws, rr, 6, f"=E{rr}*$C${L0+3}", fmt=MONEY, fill=W.CALC)
        txt(ws, rr, 7, f"=E{rr}*{LMARG}/100", fmt=MONEY, fill=W.CALC)
        txt(ws, rr, 8, f"=-E{rr}*{LCOST}/100", fmt=MONEY, fill=W.CALC)
        txt(ws, rr, 9, f"=G{rr}+H{rr}", fmt=MONEY, fill=W.FS, font=W.BOLD)
    G1 = r + len(MONTHS) - 1
    rr = G1 + 1
    txt(ws, rr, 2, "السنة كاملةً", align=W.RGT, font=W.BOLD, fill=W.FS)
    for c_ in range(3, 10):
        L = get_column_letter(c_)
        txt(ws, rr, c_,
            f"=SUM({L}{G0}:{L}{G1})" if c_ != 4 else f"=AVERAGE(D{G0}:D{G1})",
            fmt=NUM if c_ in (3, 4, 5) else MONEY, fill=W.FS, font=W.BOLD)
    ws.conditional_formatting.add(f"E{G0}:E{G1}",
                                  DataBarRule(start_type="num", start_value=0,
                                              end_type="max", color=B.ORANGE,
                                              showValue=True))
    for c_ in ("I",):
        ws.conditional_formatting.add(f"{c_}{G0}:{c_}{rr}", CellIsRule(
            operator="lessThan", formula=["0"],
            fill=PatternFill("solid", fgColor=B.T_BAD),
            font=Font(name=B.FONT, size=10, bold=True, color=B.D_BAD)))
    r = rr + 1
    r = note(ws, r,
             "المستهدف مبنيٌّ على أداء هذه المحطة وحدها: خطُّ أساسها من حجمها "
             "المقيس، وإضافتُها من فجوة ساعات ذروتها مقارنةً بشكل تصنيفها — "
             "لا من نسبة نموٍّ مفروضة على الشبكة كلها. "
             "وإن أردت مستهدفاً آخر فغيّر «الإضافة المستهدفة» وحدها.", tone="band",
             h=32)
    if s["plan"] and s["plan"]["conds"]:
        gates = [c for c in s["plan"]["conds"]
                 if c in ("فجوة مطابقة", "تغيير مسار", "بيانات دفع مفقودة")]
        if gates:
            r = note(ws, r,
                     "⛔ بوابة مفتوحة على هذه المحطة: " + " · ".join(gates) +
                     " — تُقفل قبل تحميل المستهدف على أحد، وإلا نُسبت نتيجتها "
                     "إلى الجهد وهي من البوابة.", tone="gold", h=28)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    return ws


# ═══════════════════════════════════════════════ الفهرس
def index(wb, rows):
    N = 12
    ws = wb.create_sheet("الفهرس", 0)
    W.setup(ws, [3, 26, 8, 10, 11, 12, 13, 13, 12, 12, 13, 26], freeze="A5")
    W.title(ws, "فهرس المحطات — شيتٌ لكل محطة",
            f"{len(rows)} محطة · لكلٍّ شيتٌ باسمها فيه أداؤها وحاسبة لترها "
            "ومستهدفها الشهري · والأرقام هنا ملخّصٌ منها", N)
    W.rule(ws, 3, N)
    r = 4
    W.header(ws, r, ["", "المحطة", "الرمز", "المنطقة", "نموذج العمل",
                     "مبيعات ر٢", "حجم ر٢ لتر", "صافي ر٢", "هامش الوقود",
                     "صافي اللتر بنموذجها", "مستهدف السنة لتراً", "الشيت"])
    ws.row_dimensions[r].height = 32
    r += 1
    first = r
    for s in rows:
        q = s["q2"]
        txt(ws, r, 2, s["name"], align=W.RGT, font=W.BOLD)
        txt(ws, r, 3, s["code"], font=Font(name=B.FONT, size=8, color=B.INK2))
        txt(ws, r, 4, s["region"], font=Font(name=B.FONT, size=9, color=B.INK2))
        txt(ws, r, 5, s["model"], font=Font(name=B.FONT, size=9, color=B.INK2))
        txt(ws, r, 6, q["sales2"] if q else "—", fmt=MONEY if q else None,
            fill=W.CALC)
        txt(ws, r, 7, q["vol2"] if q else "—", fmt=NUM if q else None, fill=W.CALC)
        txt(ws, r, 8, q["net2"] if q else "—", fmt=MONEY if q else None,
            fill=W.CALC)
        txt(ws, r, 9, s["fm"], fmt=PCT2, fill=W.CALC)
        net_hal = s["price_ex"] * s["fm"] * 100 - s["cpl_def"]
        txt(ws, r, 10, net_hal, fmt=HAL,
            fill=PatternFill("solid",
                             fgColor=B.T_BAD if net_hal < 0 else B.T_GOOD),
            font=Font(name=B.FONT, size=10, bold=True,
                      color=B.D_BAD if net_hal < 0 else B.D_GOOD))
        yr = sum((s["base_lpd"] + s["add_lpd"] * min(1, i / RAMP))
                 * calendar.monthrange(y, m)[1]
                 for i, (nm, y, m) in enumerate(MONTHS, 1))
        txt(ws, r, 11, yr, fmt=NUM, fill=W.CALC)
        c = ws.cell(r, 12, s["sheet"])
        c.hyperlink = f"#'{s['sheet']}'!A1"
        c.font = Font(name=B.FONT, size=9, color=B.BLUE, underline="single")
        c.alignment = W.RGT; c.border = W.BOX
        r += 1
    last = r - 1
    ws.auto_filter.ref = f"B{first-1}:L{last}"
    for col in ("G", "K"):
        ws.conditional_formatting.add(
            f"{col}{first}:{col}{last}",
            DataBarRule(start_type="num", start_value=0, end_type="max",
                        color=B.ORANGE, showValue=True))
    r += 1
    r = note(ws, r,
             "«صافي اللتر» هنا محسوبٌ بمصاريف نموذج عمل كل محطة من قائمة "
             "الدخل — استثماري ١١٫٠٨ هللة · إيجاري ٥٫٦٣ · تشغيلي ٩٫٨٤ · "
             "امتياز ٠٫٠٤ — لا برقمٍ واحد للجميع، فالفرق بينها ٢٤٦ ضعفاً. "
             "وكل هللةٍ محسوبةٌ على السعر قبل الضريبة، لأن هامش الوقود نسبةٌ "
             "من الإيراد قبلها. وداخل كل شيت تختار الأساس الذي يخصّ قرارك.",
             nc=N, h=38)
    r = note(ws, r,
             "والمصدران مختلفان في تغطيتهما: تقرير ربحية الربع الثاني يغطّي "
             "المحطات المشغّلة (٢٦)، وتقرير المالية يغطّي غيرها (٢٧) — "
             "فشرطةٌ في عمودٍ تعني أن مصدره لا يغطّي هذه المحطة، لا أن قيمتها صفر.",
             tone="gold", nc=N, h=30)
    return ws


def build():
    P, rows = load()
    rows.sort(key=lambda s: -(s["q2"]["vol2"] if s["q2"] else s["base_lpd"] * 91))
    used = set()
    for s in rows: s["sheet"] = sheet_name(s, used)
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    for s in rows: station(wb, s)
    index(wb, rows)
    wb.properties.title = "شيت كل محطة — أداء وحاسبة ومستهدف"
    wb.properties.creator = "القسم التجاري — درب"
    wb.save(OUT)
    return OUT, rows


if __name__ == "__main__":
    p, rows = build()
    nq = sum(1 for s in rows if s["q2"]); nc = sum(1 for s in rows if s["cfo"])
    print(f"كُتب {p} — {len(rows)} شيت محطة + الفهرس")
    print(f"  منها {nq} بأداء ربعيّ و{nc} بتقرير مالية")
    for s in rows[:3]:
        print(f"  {s['code']:<8}{s['name'][:18]:<20}"
              f"أساس {s['base_lpd']:>9,.0f} لتر/يوم · إضافة {s['add_lpd']:>7,.0f}"
              f" · هامش {s['fm']*100:>5.2f}٪")
