# -*- coding: utf-8 -*-
"""محافظ المحطات — محفظةٌ لكل محطة تقول كيف تبيع، بلا تعقيد

   ورقتان لا أكثر:
   ① الفهرس   — سطرٌ لكل محطة · تجد فيه محطتك وتعرف شكلها في سطر واحد
   ② المحافظ  — صفحةٌ لكل محطة · ثمانية أقسام تُقرأ بلا شرح

   ولا رقم هنا محسوبٌ باجتهاد: كلُّه مقيسٌ من كاش إن وملفات الشركة،
   وحيث ينقص المقياس تُكتب شرطةٌ لا تقدير.

   يُشغَّل من جذر المستودع:  PYTHONPATH=data python3 data/build_portfolio.py
"""
import csv, json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break
from openpyxl.formatting.rule import DataBarRule

import brand as B
import worker_model as W
import sales_plan as S

PLAN = "data/sales-plan.json"
NETS = "outlets/data/network-sales.csv"
HOURS = "outlets/data/network-hours.csv"
TARGET = "outlets/data/outlet-targeting.csv"
ANALYT = "outlets/data/analytics.csv"
UNITS = "outlets/data/units-registry.csv"
TANKS = "outlets/data/tanks.csv"
SHIFTS = "outlets/data/worker-shifts.csv"
COMPL = "outlets/data/complaints.csv"
OUT = "docs/محافظ-المحطات.xlsx"

MONEY = '#,##0;[Red](#,##0)'
NUM = '#,##0'
NUM1 = '#,##0.0'
NUM2 = '#,##0.00'
PCT = '0.0%'
PCT2 = '0.00%'

HOURN = ["منتصف الليل"] + [f"{h}" for h in range(1, 24)]


def _f(v, d=0.0):
    try: return float(str(v).replace(",", ""))
    except (TypeError, ValueError): return d


def load():
    P = json.load(open(PLAN, encoding="utf-8"))
    st = {s["code"]: dict(s) for s in P["stations"]}

    def idx(path, key, enc="utf-8-sig"):
        return {r[key]: r for r in csv.DictReader(open(path, encoding=enc))
                if r.get(key)}

    net = idx(NETS, "code", "utf-8")
    tgt = idx(TARGET, "code", "utf-8")
    ana = idx(ANALYT, "code")
    tank = idx(TANKS, "code")
    shift = idx(SHIFTS, "الرمز", "utf-8")
    hrs = idx(HOURS, "code", "utf-8")
    units = {}
    for r in csv.DictReader(open(UNITS, encoding="utf-8-sig")):
        if r.get("code"): units.setdefault(r["code"], r)
    comp = {}
    for r in csv.DictReader(open(COMPL, encoding="utf-8-sig")):
        c = r.get("station")
        if c: comp[c] = comp.get(c, 0) + 1

    for c, s in st.items():
        n, t, a = net.get(c, {}), tgt.get(c, {}), ana.get(c, {})
        rev = _f(n.get("revenue")) or s["rev"]
        s["pay"] = {k: _f(n.get("pay_" + k)) / (rev or 1)
                    for k in ("cash", "card", "fleet", "unknown")}
        s["days"] = int(_f(n.get("days"), s.get("days", 0)))
        s["spd"] = rev / (s["days"] or 1)
        vol = _f(n.get("volume")) or s["vol"]
        s["fuel"] = {"بنزين ٩١": _f(n.get("g91_vol")) / (vol or 1),
                     "بنزين ٩٥": _f(n.get("g95_vol")) / (vol or 1),
                     "ديزل": _f(n.get("diesel_vol")) / (vol or 1)}
        h = hrs.get(c)
        if h:
            v = [(_f(h["h%02d" % x]), x) for x in range(24)]
            s["top3"] = [x for _, x in sorted(v, reverse=True)[:3]]
        else:
            s["top3"] = []
        s["who"] = t.get("who", "")
        s["driver"] = t.get("driver", "")
        s["action"] = t.get("action", "")
        s["comp_n"] = int(_f(t.get("comp_n"), 0))
        s["comp_avg"] = _f(t.get("comp_avg"), 0)
        s["comp_near"] = t.get("comp_near", "")
        s["comp_who"] = t.get("comp_who", "")
        s["size"] = a.get("size", "")
        s["growth"] = _f(a.get("growth"), 0)
        s["campaign"] = a.get("campaign", "")
        sh = shift.get(c, {})
        s["staff"] = int(_f(sh.get("عدد العمّال"), 0))
        s["eve_share"] = _f(sh.get("حصة الوردية المسائية"), 0)
        u = units.get(c, {})
        s["u_total"] = int(_f(u.get("total_n"), 0))
        s["u_leased"] = int(_f(u.get("total_leased"), 0))
        s["u_vacant"] = int(_f(u.get("total_vacant"), 0))
        s["u_shop"] = int(_f(u.get("shop_n"), 0))
        s["u_shop_v"] = int(_f(u.get("shop_vacant"), 0))
        s["u_cat"] = u.get("category", "")
        tk = tank.get(c, {})
        cap = sum(_f(tk.get(k)) for k in ("cap_91", "cap_95", "cap_98", "cap_diesel"))
        s["tank_days"] = cap / (s["lpd"] or 1) if cap else 0
        s["complaints"] = comp.get(c, 0)
        s["gate"] = [x for x in s["conds"]
                     if x in ("فجوة مطابقة", "تغيير مسار", "بيانات دفع مفقودة")]
    return P, sorted(st.values(), key=lambda s: (s["region"], -s["lpd"]))


# ── الوصف في سطر: من الأرقام المقيسة لا من الاسم
def shape(s):
    p = []
    v = s["vpd"]
    p.append("حركة عالية" if v >= 1500 else
             "حركة متوسطة" if v >= 600 else "حركة منخفضة")
    L = s["lpv"]
    p.append("سلّة كبيرة" if L >= 35 else
             "سلّة متوسطة" if L >= 25 else "سلّة صغيرة")
    d = s["diesel"]
    p.append("ديزل غالب" if d >= 0.40 else
             "ديزل وبنزين" if d >= 0.20 else "بنزين")
    if s["pay"]["fleet"] >= 0.10: p.append("أسطول متعاقَد")
    elif s["pay"]["cash"] >= 0.55: p.append("نقدي غالب")
    else: p.append("نقدي وبطاقة")
    if s["night"] >= 0.55: p.append("ليلية")
    return " · ".join(p)


def story(s):
    """كيف تبيع هذه المحطة — بالكلمات، من الأرقام المقيسة وحدها"""
    d = s["diesel"]; pet = 1 - d
    freq = ("حركتها عالية" if s["vpd"] >= 1500 else
            "حركتها متوسطة" if s["vpd"] >= 600 else "حركتها منخفضة")
    bask = ("وسلّتها كبيرة" if s["lpv"] >= 35 else
            "وسلّتها متوسطة" if s["lpv"] >= 25 else "وسلّتها صغيرة")
    a = (f"تبيع {s['vpd']:,.0f} زيارة في اليوم بمتوسط {s['lpv']:.1f} لتراً "
         f"للزيارة — {freq} {bask}.")
    if d >= 0.40:
        b = f"الديزل غالبُ لتراتها ({d*100:.0f}٪)، فسوقها شاحنات قبل أن يكون سيارات."
    elif d >= 0.20:
        b = (f"تبيع بنزيناً وديزلاً معاً ({pet*100:.0f}٪ بنزين و{d*100:.0f}٪ ديزل) — "
             "أي طلبان في موقع واحد.")
    else:
        b = f"بنزينٌ في أغلبها ({pet*100:.0f}٪)، فسوقها سيارات الأفراد."
    c = f"ذروتها الساعة {s['peak']}"
    n = s["night"]
    if n >= 0.55: c += f"، وأكثر من نصف زياراتها ليلاً ({n*100:.0f}٪)"
    elif n >= 0.48: c += f"، ونحو نصف زياراتها ليلاً ({n*100:.0f}٪)"
    else: c += f"، وليلها {n*100:.0f}٪ من زياراتها"
    if s["staff"]:
        # تمييز العدد: ٣–١٠ جمعٌ مجرور، وما فوقها مفردٌ منصوب
        w = "عمّال" if 3 <= s["staff"] <= 10 else "عاملاً"
        c += f"، ويخدمها {s['staff']} {w}"
    c += "."
    if s["pay"]["unknown"] > 0.5:
        e = "ولا نعرف كيف يدفع زبائنها — وسيلة الدفع غير مسجَّلة في أغلب مبيعاتها."
    elif s["pay"]["fleet"] >= 0.10:
        e = (f"ويدفع {s['pay']['fleet']*100:.0f}٪ من مبيعاتها بحسابات أسطول "
             "متعاقَدة — وهي نسبةٌ تُبنى عليها.")
    elif s["pay"]["cash"] >= 0.55:
        e = (f"ويدفع زبائنها نقداً في أغلبهم ({s['pay']['cash']*100:.0f}٪) — "
             "وهو ما يجعل الرقابة على الصندوق أهمّ من أي حملة.")
    else:
        e = (f"ويدفعون نقداً وبطاقةً بالتناصف تقريباً "
             f"({s['pay']['cash']*100:.0f}٪ و{s['pay']['card']*100:.0f}٪).")
    return " ".join([a, b, c, e])


def diesel_read(s):
    """هل ديزلها متعاقَدٌ أم نقدي؟ — القناة تُعرف بالحساب لا بالاسم"""
    d, f, u = s["diesel"], s["pay"]["fleet"], s["pay"]["unknown"]
    num = f"الديزل {d*100:.0f}٪ من اللترات والأسطول المسجَّل {f*100:.1f}٪ من المبيعات"
    if u > 0.5:
        return f"{num} — ووسيلة الدفع غير مسجَّلة في أغلب المبيعات، فلا يُحكم بعد"
    if d < 0.20:
        return f"{num} — فديزلها محدود والأسطول ليس سوقها"
    if f >= d * 0.6:
        return f"{num} — فأغلب ديزلها متعاقَد"
    if f >= d * 0.25:
        return f"{num} — فجزءٌ من ديزلها متعاقَد والباقي نقدي"
    return (f"{num} — فديزلها نقديٌّ في أغلبه: شاحنات تمرّ ولا تُسجَّل، "
            "وهنا فرصةُ تعاقدٍ لا حملة")


# ═══════════════════════════════════════════════ ① الفهرس
def index(wb, rows):
    NC = 13
    ws = wb.create_sheet("الفهرس")
    W.setup(ws, [3, 24, 8, 9, 9, 10, 11, 9, 9, 9, 11, 34, 11], freeze="A5")
    W.title(ws, "فهرس المحطات — شكل كل محطة في سطر واحد",
            f"{len(rows)} محطة مقيسة من كاش إن · مرتَّبة بالمنطقة ثم بالحجم · "
            "وصفحة كل محطة في ورقة «المحافظ»", NC)
    W.rule(ws, 3, NC)
    r = 4
    W.header(ws, r, ["", "المحطة", "الرمز", "المنطقة", "التصنيف", "زيارة/يوم",
                     "لتر/يوم", "لتر/زيارة", "ديزل٪", "نقدي٪", "ساعة الذروة",
                     "كيف تبيع", "صفحتها"])
    ws.row_dimensions[r].height = 30
    r += 1
    first = r
    for s in rows:
        gate = bool(s["gate"])
        fill = PatternFill("solid", fgColor=B.T_GOLD) if gate else None
        c = ws.cell(r, 2, s["name"]); c.font = W.BOLD; c.alignment = W.RGT
        c.border = W.BOX
        if fill: c.fill = fill
        # النقدي لا يُعرض حين تكون الوسيلة غير مسجَّلة — شرطةٌ لا صفر
        known = s["pay"]["unknown"] <= 0.5
        vals = [(3, s["code"], None), (4, s["region"], None), (5, s["seg"], None),
                (6, s["vpd"], NUM), (7, s["lpd"], NUM), (8, s["lpv"], NUM2),
                (9, s["diesel"], PCT),
                (10, s["pay"]["cash"] if known else "—", PCT if known else None),
                (11, f"{int(s['peak']):02d}:٠٠", None)]
        for col, v, fm in vals:
            x = ws.cell(r, col, v); x.border = W.BOX; x.alignment = W.CTR
            x.font = Font(name=B.FONT, size=9, color=B.INK2) if col in (3, 4, 5) \
                else W.BLACK
            if fm: x.number_format = fm
            x.fill = fill or W.CALC
        x = ws.cell(r, 12, ("⛔ " if gate else "") + shape(s))
        x.border = W.BOX; x.alignment = W.RGT
        x.font = Font(name=B.FONT, size=9, bold=gate,
                      color=B.D_GOLD if gate else B.INK)
        if fill: x.fill = fill
        x = ws.cell(r, 13, s["page"]); x.border = W.BOX; x.alignment = W.CTR
        x.font = Font(name=B.FONT, size=9, color=B.BLUE)
        if fill: x.fill = fill
        r += 1
    last = r - 1
    for col in ("F", "G"):
        ws.conditional_formatting.add(
            f"{col}{first}:{col}{last}",
            DataBarRule(start_type="num", start_value=0, end_type="max",
                        color=B.ORANGE, showValue=True))
    ws.auto_filter.ref = f"B{first-1}:M{last}"
    ws.merge_cells(start_row=r + 1, start_column=1, end_row=r + 1, end_column=NC)
    c = ws.cell(r + 1, 1,
                "⛔ الصفّ الذهبي = بوابة مفتوحة على هذه المحطة (فجوة مطابقة أو "
                "تغيير مسار أو بيانات دفع مفقودة) — تُقفل قبل أي حملة، "
                "وتفصيلها في صفحة المحطة.")
    c.font = Font(name=B.FONT, size=9, bold=True, color=B.D_GOLD)
    c.fill = PatternFill("solid", fgColor=B.T_GOLD)
    c.alignment = W.WRAP; c.border = W.BOX
    ws.row_dimensions[r + 1].height = 28
    return ws


# ═══════════════════════════════════════════════ ② المحافظ
NC = 5          # فراغ · العنوان · القيمة · الوحدة · الملاحظة


def line(ws, r, label, value, *, fmt=None, unit="", tone=None, big=False,
         note=""):
    """سطرٌ واحد: عنوانٌ ثم رقمٌ ثم وحدته ثم ملاحظةٌ قصيرة إن لزمت"""
    c = ws.cell(r, 2, label); c.font = W.BOLD; c.alignment = W.RGT; c.border = W.BOX
    fill = PatternFill("solid", fgColor={
        "good": B.T_GOOD, "bad": B.T_BAD, "gold": B.T_GOLD}.get(tone, B.T_NEUTRAL))
    col = {"good": B.D_GOOD, "bad": B.D_BAD, "gold": B.D_GOLD}.get(tone, B.INK)
    x = ws.cell(r, 3, value)
    x.font = Font(name=B.FONT, size=13 if big else 11, bold=True, color=col)
    x.alignment = W.CTR; x.fill = fill; x.border = W.BOX
    if fmt: x.number_format = fmt
    u = ws.cell(r, 4, unit)
    u.font = Font(name=B.FONT, size=9, color=B.INK3)
    u.alignment = W.RGT; u.border = W.BOX; u.fill = fill
    n = ws.cell(r, 5, note)
    n.font = Font(name=B.FONT, size=9, color=B.INK3)
    n.alignment = W.RGT; n.border = W.BOX
    return r + 1


def wide(ws, r, label, text, *, tone=None, h=None):
    """نصٌّ يمتدّ: بعنوانٍ فيبدأ من عمود القيمة، وبلا عنوانٍ فيمتدّ كله"""
    c1 = 2
    if label is not None:
        c = ws.cell(r, 2, label); c.font = W.BOLD; c.alignment = W.RGT
        c.border = W.BOX
        c1 = 3
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=NC)
    fill = PatternFill("solid", fgColor={
        "good": B.T_GOOD, "bad": B.T_BAD, "gold": B.T_GOLD,
        "band": B.T_ORANGE}.get(tone, B.CARD))
    col = {"good": B.D_GOOD, "bad": B.D_BAD, "gold": B.D_GOLD,
           "band": B.ORANGE}.get(tone, B.INK2)
    x = ws.cell(r, c1, text)
    x.font = Font(name=B.FONT, size=10, bold=tone is not None, color=col)
    x.alignment = W.WRAP; x.fill = fill; x.border = W.BOX
    for j in range(c1 + 1, NC + 1):
        cc = ws.cell(r, j); cc.border = W.BOX; cc.fill = fill
    if h: ws.row_dimensions[r].height = h
    return r + 1


def sect(ws, r, text):
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=NC)
    c = ws.cell(r, 2, text)
    c.font = Font(name=B.FONT, size=11, bold=True, color=B.ORANGE)
    c.fill = PatternFill("solid", fgColor=B.T_ORANGE)
    c.alignment = W.RGT; c.border = W.BOX
    for j in range(3, NC + 1):
        cc = ws.cell(r, j); cc.border = W.BOX
        cc.fill = PatternFill("solid", fgColor=B.T_ORANGE)
    ws.row_dimensions[r].height = 20
    return r + 1


def card(ws, r, s):
    """صفحةُ محطةٍ واحدة — ثمانية أقسام تُقرأ بلا شرح"""
    # العنوان
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=NC)
    c = ws.cell(r, 2, f"{s['name']}  ·  {s['code']}")
    c.font = Font(name=B.FONT, size=18, bold=True, color=B.BGRAY)
    c.alignment = W.RGT; c.border = W.BOX
    ws.row_dimensions[r].height = 30
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=NC)
    c = ws.cell(r, 2, f"{s['region']}  ·  {s['seg']}  ·  {s['size']}  ·  "
                      f"{s['days']} يوماً مقيساً  ·  {shape(s)}")
    c.font = Font(name=B.FONT, size=10, color=B.INK2)
    c.alignment = W.RGT; c.border = W.BOX
    ws.row_dimensions[r].height = 18
    r += 1
    for j in range(2, NC + 1):
        ws.cell(r, j).fill = PatternFill("solid", fgColor=B.ORANGE)
    ws.row_dimensions[r].height = 3.5
    r += 1
    r = wide(ws, r, None, story(s), tone="band", h=46)

    # ① كم تبيع
    r = sect(ws, r, "① كم تبيع")
    r = line(ws, r, "زيارة في اليوم", round(s["vpd"]), fmt=NUM, unit="زيارة", big=True)
    r = line(ws, r, "لتر في اليوم", round(s["lpd"]), fmt=NUM, unit="لتر", big=True)
    r = line(ws, r, "ريال في اليوم", round(s["spd"]), fmt=MONEY, unit="ريال", big=True)
    r = line(ws, r, "لتر لكل زيارة", s["lpv"], fmt=NUM2, unit="لتر")
    r = line(ws, r, "متوسط الفاتورة", s["inv"], fmt=NUM2, unit="ريال")
    g = s["growth"]
    r = line(ws, r, "النمو المسجَّل", g / 100, fmt='+0.0%;-0.0%', unit="ر١ مقابل ر٢",
             tone="good" if g >= 0 else "bad")
    r = wide(ws, r, None,
             "والنمو هنا مقارنةُ ربعٍ بربع — فإن كان الربع المقارَن موسماً فالرقم "
             "يقول عن الموسم أكثر مما يقول عن المحطة.", h=26)

    # ② ماذا تبيع
    r = sect(ws, r, "② ماذا تبيع")
    for k in ("بنزين ٩١", "بنزين ٩٥", "ديزل"):
        r = line(ws, r, k, s["fuel"][k], fmt=PCT, unit="من اللترات")
    r = wide(ws, r, None, diesel_read(s), tone="gold", h=26)

    # ③ لمن تبيع
    r = sect(ws, r, "③ لمن تبيع")
    d = S.SEGDEF.get(s["seg"], ("", "", ""))
    r = wide(ws, r, "التصنيف", f"{s['seg']} — {d[0]} · {d[2]}", h=24)
    r = line(ws, r, "الأسطول المسجَّل", s["pay"]["fleet"], fmt=PCT2,
             unit="من المبيعات")
    r = line(ws, r, "حصة الليل", s["night"], fmt=PCT, unit="من الزيارات")
    if s["comp_n"]:
        w = "محطات" if 3 <= s["comp_n"] <= 10 else "محطة"
        r = wide(ws, r, "يزاحمها",
                 f"{s['comp_n']} {w} حولها · متوسط تقييمها {s['comp_avg']:.2f}★ · "
                 f"أقربها {s['comp_who']} على {s['comp_near']}", h=24)
    else:
        r = wide(ws, r, "يزاحمها", "لم تُمسح محيطها بعد — يلزم مسحٌ ميداني", h=24)

    # ④ متى تبيع
    r = sect(ws, r, "④ متى تبيع")
    r = line(ws, r, "ساعة الذروة", s["peak"], fmt=NUM, unit="من ٢٤")
    if s["top3"]:
        r = wide(ws, r, "أكثر ثلاث ساعات",
                 " · ".join(f"{h:02d}:٠٠" for h in s["top3"]), h=20)
    have_sh = s["staff"] > 0 or s["eve_share"] > 0
    r = line(ws, r, "الوردية المسائية ١٢–٢٣",
             s["eve_share"] if have_sh else "—", fmt=PCT if have_sh else None,
             unit="من المعاملات",
             note="" if have_sh else "لا بيان ورديات لهذه المحطة")
    r = line(ws, r, "عدد العمّال", s["staff"] or "—",
             fmt=NUM if s["staff"] else None, unit="عامل",
             note="" if s["staff"] else "لا بيان عمّال — يلزم من العمليات")
    if s["staff"]:
        r = wide(ws, r, None,
                 f"أي نحو {s['vpd']/s['staff']:,.0f} زيارة لكل عامل في اليوم — "
                 "موزَّعةً على وردیتين، والمساء يحمل أكثرها.", h=24)

    # ⑤ كيف يدفعون
    r = sect(ws, r, "⑤ كيف يدفعون")
    for k, lab in (("cash", "نقدي"), ("card", "بطاقة"), ("fleet", "أسطول"),
                   ("unknown", "غير مسجَّلة")):
        bad = k == "unknown" and s["pay"][k] > 0.02
        r = line(ws, r, lab, s["pay"][k], fmt=PCT2, unit="من المبيعات",
                 tone="bad" if bad else None)
    if s["pay"]["unknown"] > 0.5:
        r = wide(ws, r, None,
                 "⛔ أغلب مبيعاتها غير مسجَّلة الوسيلة — فالنسب أعلاه لا تُقرأ "
                 "حصصاً بل دليلاً على أن النظام لم يُفعَّل. يُضبط أولاً، وإلا "
                 "نُسب أثرُ التسجيل إلى الحملة.", tone="gold", h=32)
    elif s["pay"]["unknown"] > 0.02:
        r = wide(ws, r, None,
                 "⛔ وسيلة الدفع غير مسجَّلة في جزءٍ معتبَر من المبيعات — "
                 "يُضبط النظام قبل أي حملة، وإلا نُسب أثرُ التسجيل إلى الحملة.",
                 tone="gold", h=28)

    # ⑥ وحداتها التجارية
    r = sect(ws, r, "⑥ وحداتها التجارية")
    if s["u_total"]:
        r = line(ws, r, "الوحدات", s["u_total"], fmt=NUM, unit="وحدة")
        r = line(ws, r, "المؤجَّرة", s["u_leased"], fmt=NUM, unit="وحدة")
        r = line(ws, r, "الشاغرة", s["u_vacant"], fmt=NUM, unit="وحدة",
                 tone="bad" if s["u_vacant"] >= 10 else None)
        per = s["vpd"] / s["u_vacant"] if s["u_vacant"] else 0
        if not s["u_vacant"]:
            r = wide(ws, r, None,
                     "مؤجَّرة بالكامل — لا فرصة تأجير هنا اليوم", tone="good",
                     h=20)
        if s["u_vacant"]:
            r = wide(ws, r, None,
                     f"{per:,.0f} زيارة في اليوم لكل وحدة شاغرة — "
                     + ("وهو فوق وسيط الشبكة (١١١) — فالحركة موجودة فسعِّر منها"
                        if per >= 111 else
                        "وهو دون وسيط الشبكة (١١١) — فلا يُصلحها خفض إيجار"), h=26)
    else:
        r = wide(ws, r, "الوحدات", "لا سجلّ وحدات لهذه المحطة في ملفاتنا", h=20)

    # ⑦ الفرصة
    r = sect(ws, r, "⑦ الفرصة — بالأرقام")
    r = line(ws, r, "فجوة ساعات الذروة", float(s["gap_peak"]), fmt=NUM2,
             unit="معاملة/يوم", tone="bad" if s["gap_peak"] > 20 else None)
    r = line(ws, r, "قيمتها", round(s["sar_txn"]), fmt=MONEY, unit="ريال/سنة",
             tone="good", big=True)
    r = line(ws, r, "فجوة السلة", float(s["gap_fill"]), fmt=NUM2,
             unit="لتر/زيارة")
    r = line(ws, r, "قيمتها", round(s["sar_fill"]), fmt=MONEY, unit="ريال/سنة")
    r = wide(ws, r, None,
             "فجوة الذروة تُطالَب بها الوردية — أمّا فجوة السلة فتُفسَّر بمزيج "
             "الوقود ولا يملكها عامل.", h=24)

    # ⑧ الفعل
    r = sect(ws, r, "⑧ الفعل")
    if s["gate"]:
        r = wide(ws, r, "⛔ قبل أي حملة",
                 " · ".join(s["gate"]) + " — تُقفل أولاً، وإلا نُسبت نتيجتها "
                 "إلى الحملة وهي من البوابة.", tone="gold", h=28)
    r = wide(ws, r, "المحرّك الأول", s["driver"] or "—", tone="band", h=20)
    r = wide(ws, r, "الفعل", s["action"] or "—", h=26)
    extra = []
    if s["tank_days"] and s["tank_days"] < 4:
        extra.append(f"تغطية الخزان {s['tank_days']:.1f} يوم — دون أربعة أيام، "
                     "فيُثبَّت التوريد قبل رفع الطلب")
    if s["complaints"] >= 5:
        extra.append(f"{s['complaints']} شكوى مسجَّلة — أكثرها بطء خدمة")
    if s["campaign"] == "نعم":
        extra.append("نُفِّذت عليها حملة — يلزم قياسها بنافذة قبل/أثناء/بعد")
    r = wide(ws, r, "ما ينبغي الانتباه له",
             " · ".join(extra) if extra else "لا شيء إضافي مرصود", h=28)
    return r


def portfolios(wb, rows):
    ws = wb.create_sheet("المحافظ")
    W.setup(ws, [3, 28, 15, 16, 44], freeze=None)
    ws.freeze_panes = None
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    r = 1
    for i, s in enumerate(rows):
        s["page"] = f"صف {r}"
        r = card(ws, r, s)
        r += 2
        if i < len(rows) - 1:
            ws.row_breaks.append(Break(id=r - 1))
    return ws


def build():
    P, rows = load()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    portfolios(wb, rows)          # تُبنى أولاً لتُعرف أرقام الصفوف
    index(wb, rows)
    wb.move_sheet("الفهرس", offset=-1)
    wb.properties.title = "محافظ المحطات — درب"
    wb.properties.creator = "القسم التجاري — درب"
    wb.save(OUT)
    return OUT, rows


if __name__ == "__main__":
    p, rows = build()
    print(f"كُتب {p} — {len(rows)} محفظة")
    for s in rows[:3]:
        print(f"  {s['code']:<7}{s['name'][:20]:<22}{shape(s)}")
