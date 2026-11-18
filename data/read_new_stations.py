# -*- coding: utf-8 -*-
"""قراءةُ ملف الموازنة واستخلاصُ مُدخَلاتِ توقّعِ المحطات الجديدة — مقيسةً لا مقدَّرة

   المصدر: «Sales & Rent Budget 2027 - LV …xlsx» — ثلاثُ أوراقٍ منه:
     · `New Stations 2027`           — الـ٣٨ محطةً وتواريخُ افتتاحها
     · `L&V 2027 vs 2026`            — لترُ كلِّ محطةٍ ومنتجٍ شهراً بشهر للسنتين
     · `GC│Liters & Value subledger` — سماتُ المحطة: المنطقة والمدينة والنموذج

   وثلاثةُ مطبّاتٍ في الملف تُعالَج هنا صراحةً:

     ① كودُ `L&V` يحمل لاحقةَ المنتج (`MK12791` = المحطة `MK127` + المنتج `91`)
        فتُجمَع المحطةُ بأول خمسةِ محارفَ من الكود، وإلا صارت كلُّ محطةٍ ثلاثاً.

     ② أكتوبر–ديسمبر ٢٠٢٦ تقديرٌ لا فعليّ (قِيَمٌ متساويةٌ متتابعة)، فكلُّ
        قياسٍ هنا يقتصر على يناير–سبتمبر.

     ③ «الناضجة» تعني حجماً في كلِّ شهرٍ من ينا–سبت ٢٠٢٦؛ ومن دون هذا الشرط
        تلوّثُ محطةٌ افتُتحت في أثناء ٢٠٢٦ وسيطَ النظراء ومنحنى الموسم معاً.

   ويُستخلَص خمسةُ مُدخَلات:
     · معدّلُ النظير   — وسيطُ لتر/يوم لمحطاتِ المدينة (ن≥٣) وإلا المنطقة،
                         أو وسيطُ فئةٍ مقيسةٍ حين تكون المحطةُ من غيرِ جنسِ جيرانها
     · منحنى التدرّج   — من المحطات التي افتُتحت في أثناء ٢٠٢٦ ولها ٥ أشهرٍ فعليّة
     · كسرُ شهرِ الافتتاح — أيامُ التشغيل في شهرِ الافتتاح ÷ أيامِ الشهر
     · معاملُ الموسم   — منحنياتٌ مقيسةٌ يُختار منها واحدٌ لكلِّ محطة
     · مزيجُ المنتجات  — حصصُ ٩١/٩٥/الديزل الفعليّةُ لمنطقةِ المحطة أو لفئتها

   والفئاتُ المقيسة — تُنتزع من الناضجة نفسِها لا من تسميةٍ في الملف:
     · «ممرُّ شاحنات»  حصّةُ ديزلٍ ٤٠–٥٥٪  — للمحطة ذاتِ مسارِ الشاحنات
     · «ديزلٌ غالب»    حصّةُ ديزلٍ فوق ٥٥٪ — للمحطة التي يغلب عليها النقل
   ومقاسُ الملف (`Small` / `Medium`) لا يُعوَّل عليه: كلُّ الـ٣٨ مسجَّلةٌ
   `Small Gas Station` ومنها ما مساحتُه فوق مئةِ ألفِ متر، ووسيطُ الـ`Medium`
   القائمة ٣١٬١١٣ ل/يوم بتشتُّتٍ من ١٥٬٦٦٢ إلى ٩٦٬٨٨٠ — فالمقاسُ لا يفرز.

   OVERRIDE — ما تعرفه الإدارةُ عن الموقعِ ولا يظهر في الملف. يُكتب هنا
   بمصدره صريحاً، لأنّه مُدخَلٌ إداريٌّ لا قياسٌ مستخرَج.

   يُشغَّل:  python3.13 -I data/read_new_stations.py <مسار الملف>
             (و‏-I لأنّ الملف مُرفَقٌ غيرُ موثوق)
"""
import collections
import csv
import datetime as dt
import io
import json
import math
import statistics as st
import sys

import openpyxl

OUT = "outlets/data/new-stations-2027.json"
UNITS = "outlets/data/units-registry.csv"      # سجلُّ الوحدات — ١٨٦ محطة
LEASE = "outlets/data/leasing-own.csv"         # أسماءُ المستأجرين بالمحطة
DAYS = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]   # ٢٠٢٧ — فبراير ٢٨
NACT = 9          # ينا–سبت: آخرُ شهرٍ فعليٍّ في ملف ٢٠٢٦
PRODS = ("91", "95", "diesel")

# علاماتٌ كبرى — لاختبارِ هل تُفسّر العلامةُ اللترَ أم عددُ الوحدات
BRANDS = ["ماكدونالدز", "كرسبي كريم", "تيم هورتنز", "دانكن", "بارنز",
          "سينابون", "بابا", "هرفي", "كودو", "البيك", "دومينوز", "ستاربكس",
          "كنتاكي", "شاورمر", "بيتزا هت", "صب واي", "تشيليز", "نستو",
          "بنده", "كارفور", "العثيم", "فريش واي"]

# كسرُ شهرِ الافتتاح في عيّنةِ القياس — تواريخُ افتتاحِ الـ٢٢ غيرُ مسجَّلةٍ في
# الملف، فيُفترَض توزُّعُها المنتظمُ داخلَ الشهر ⇒ نصفُ شهرٍ وسطياً. وهو
# فرضٌ لا قياس، ولذلك هو خليّةٌ تُغيَّر في ورقةِ المُدخَلات.
FRAC_SAMPLE = 0.50

# فئاتُ الديزل — حدودُ حصّةِ الديزل التي تُنتزع بها الفئةُ من الناضجة
CLASSES = [("ممرُّ شاحنات", 0.40, 0.55),
           ("ديزلٌ غالب", 0.55, 1.01)]

# ما تعرفه الإدارةُ عن الموقعِ ولا يَظهر في الملف — مُدخَلٌ إداريٌّ مُعلَن
OVERRIDE = {
    "MK044": dict(cls="ممرُّ شاحنات", basis="من عددِ الوحدات",
                  why="مساحةٌ فوق ١٠٠ ألف م² · مسارُ ديزلٍ مخصَّصٌ للشاحنات · "
                      "وتجارةٌ كثيفةٌ بعلاماتٍ كبرى تجرّ حركةً للموقع "
                      "— الإدارة التجارية"),
    "MK045": dict(cls="ممرُّ شاحنات", prof="صيفيٌّ مقيس",
                  basis="من عددِ الوحدات",
                  why="مساحةٌ فوق ١٠٠ ألف م² · مسارُ ديزلٍ مخصَّصٌ للشاحنات · "
                      "تجارةٌ كثيفةٌ بعلاماتٍ كبرى · وموسمٌ صيفيٌّ بافتتاحِ "
                      "ريكسوس ومنتجعاتٍ قبلها — الإدارة التجارية"),
}

num = lambda x: float(x) if isinstance(x, (int, float)) else 0.0


def read(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)

    # ① سماتُ المحطات — أولُ ظهورٍ لكلِّ كودٍ مختصرٍ يكفي
    attr = {}
    for r in wb["GC│Liters & Value subledger"].iter_rows(min_row=8, max_col=12,
                                                         values_only=True):
        if not r[10] or not r[7]:
            continue
        attr.setdefault(str(r[10]).strip()[:5], dict(
            name=str(r[7]).split(" - ")[0].strip(),
            model=str(r[2]).strip().title(),
            reg=str(r[3]).strip().replace(" Region", ""),
            city=str(r[4]).strip()))

    # ② الـ٣٨ محطةً الجديدة
    new = {}
    for r in wb["New Stations 2027"].iter_rows(min_row=2, max_col=10,
                                               values_only=True):
        if not r[8] or not isinstance(r[7], (dt.date, dt.datetime)):
            continue
        d = r[7].date() if isinstance(r[7], dt.datetime) else r[7]
        new[str(r[8]).strip()] = dict(
            name=str(r[6]).split(" - ")[0].strip(),
            model=str(r[1]).strip().title(),
            reg=str(r[2]).strip().replace(" Region", ""),
            city=str(r[3]).strip().title(),
            open=d.isoformat())

    # ③ اللترات — محطةً ومنتجاً وشهراً
    vol = collections.defaultdict(
        lambda: collections.defaultdict(lambda: {2026: [0.0] * 12,
                                                 2027: [0.0] * 12}))
    for r in wb["L&V 2027 vs 2026"].iter_rows(min_row=5, max_col=18,
                                              values_only=True):
        if not r[0] or str(r[4]).strip() not in ("2026", "2027"):
            continue
        code = str(r[0]).strip()[:5]          # ← المطبُّ ①
        p = str(r[3]).strip().lower()
        if p not in PRODS:
            continue
        y = int(str(r[4]).strip())
        for i, x in enumerate(r[5:17]):
            vol[code][p][y][i] += num(x)
    return attr, new, {k: dict(v) for k, v in vol.items()}


def totals(v, year):
    return [sum(v[p][year][i] for p in v) for i in range(12)]


def measure_ramp(vol, new):
    """منحنى التدرّج من افتتاحاتِ ٢٠٢٦ الفعليّة

       لكلِّ محطةٍ أولُ شهرِ حجمٍ بعد يناير وبلا فجوات، ومستوى الاستقرار
       متوسّطُ آخرِ شهرين فعليَّين، ثم وسيطُ النسب لكلِّ شهرِ تشغيل.
    """
    ramp = collections.defaultdict(list)
    n = 0
    for c, v in vol.items():
        if c in new:
            continue
        a = totals(v, 2026)[:NACT]
        nz = [i for i, x in enumerate(a) if x > 0]
        if not nz or nz[0] == 0 or len(nz) != NACT - nz[0]:
            continue                       # تعمل من يناير، أو فيها فجوة
        o = nz[0]
        if NACT - o < 5:
            continue                       # أقلُّ من خمسةِ أشهرٍ فعليّة
        rt = [a[o + k] / DAYS[o + k] for k in range(NACT - o)]
        plat = st.mean(rt[-2:])
        if plat <= 0:
            continue
        n += 1
        for k, x in enumerate(rt, 1):
            ramp[k].append(x / plat)
    curve = [round(st.median(ramp[k]), 3) for k in range(1, 8) if len(ramp[k]) >= 5]
    while len(curve) < 7:
        curve.append(1.0)
    return curve, n


def measure_peers(vol, new, attr):
    """وسيطُ لتر/يوم للناضجة — بالمدينة وبالمنطقة وبالشبكة"""
    rate = {}
    for c, v in vol.items():
        if c in new or c not in attr:
            continue
        a = totals(v, 2026)
        if not all(x > 0 for x in a[:NACT]):
            continue                       # ← المطبُّ ③
        rate[c] = st.mean([a[i] / DAYS[i] for i in range(NACT)])
    city, reg = collections.defaultdict(list), collections.defaultdict(list)
    for c, x in rate.items():
        city[attr[c]["city"].strip().lower()].append(x)
        reg[attr[c]["reg"]].append(x)
    return rate, city, reg


def measure_classes(vol, new, attr):
    """فئاتٌ مُنتزعةٌ من الناضجة بحصّةِ الديزل — لا بتسميةِ المقاس في الملف

       فالمقاسُ المسجَّل لا يفرز: كلُّ الـ٣٨ `Small` ومنها ما يفوق مئةَ ألفِ
       متر. وحصّةُ الديزل تفرز، لأنّها أثرُ الشاحناتِ في السلّة لا اسمُها.
    """
    out = {}
    for name, lo, hi in CLASSES:
        rates, mix = [], collections.Counter()
        members = []
        for c, v in vol.items():
            if c in new or c not in attr:
                continue
            t = totals(v, 2026)
            if not all(x > 0 for x in t[:NACT]):
                continue
            s = sum(t[:NACT])
            dsl = sum(v["diesel"][2026][:NACT]) if "diesel" in v else 0.0
            share = dsl / s if s else 0
            if not lo <= share < hi:
                continue
            rt = st.mean([t[i] / DAYS[i] for i in range(NACT)])
            rates.append(rt)
            members.append([attr[c]["name"], round(rt), round(share, 3)])
            for p in PRODS:
                mix[p] += sum(v[p][2026][:NACT]) if p in v else 0.0
        if len(rates) < 3:
            continue
        sm = sum(mix[p] for p in PRODS)
        rates.sort()
        out[name] = dict(rate=round(st.median(rates)), n=len(rates),
                         q3=round(rates[3 * len(rates) // 4]), top=round(rates[-1]),
                         mix={p: round(mix[p] / sm, 4) for p in PRODS},
                         members=sorted(members, key=lambda x: -x[1]))
    return out


def measure_profiles(vol, new, attr):
    """منحنياتُ موسمٍ مقيسةٌ يُختار منها واحدٌ لكلِّ محطة

       «صيفيٌّ مقيس» — أعلى عشرِ محطاتٍ في نسبةِ (يون–أغس ÷ ينا–ماي+سبت).
       وأشهرُ أكت–ديس لا فعليَّ لها في ٢٠٢٦، فتُوضع على مستوى ينا–أبر
       المقيس — وهو فرضٌ مُعلَنٌ لا قياس.
    """
    cand = []
    for c, v in vol.items():
        if c in new or c not in attr:
            continue
        t = totals(v, 2026)
        if not all(x > 0 for x in t[:NACT]):
            continue
        d = [t[i] / DAYS[i] for i in range(NACT)]
        base = st.mean([d[i] for i in (0, 1, 2, 3, 4, 8)])
        if base <= 0:
            continue
        mono = all(d[i] <= d[i + 1] * 1.02 for i in range(NACT - 1))
        cand.append((c, st.mean(d[5:8]) / base, mono))
    free = sorted([x for x in cand if not x[2]], key=lambda x: -x[1])
    top = [c for c, _, _ in free[:10]]
    agg = [sum(totals(vol[c], 2026)[i] for c in top) / DAYS[i]
           for i in range(NACT)]
    b = st.mean(agg)
    summer = [round(agg[i] / b, 3) for i in range(NACT)]
    off = round(st.mean(summer[:4]), 3)        # مستوى ينا–أبر المقيس
    return dict(
        summer=summer + [off] * 3,
        summer_n=len(top),
        summer_median=round(st.median([x[1] for x in cand]), 3),
        summer_members=[[attr[c]["name"], attr[c]["city"], round(r, 2)]
                        for c, r, _ in free[:10]])


def read_units():
    """سجلُّ الوحدات وأسماءُ المستأجرين — عددُ الوحداتِ المبنيّةِ والمؤجَّرة"""
    U, Lz = {}, {}
    for r in csv.DictReader(io.open(UNITS, encoding="utf-8")):
        k = (r.get("code") or "").strip()
        if not k:
            continue
        g = lambda f: int(float(r.get(f) or 0))
        U[k] = dict(built=g("total_n"), leased=g("total_leased"),
                    vacant=g("total_vacant"), shop=g("shop_n"))
    for r in csv.DictReader(io.open(LEASE, encoding="utf-8")):
        k = (r.get("code") or r.get("﻿code") or "").strip()
        if not k:
            continue
        txt = " ".join((r.get(f) or "") for f in
                       ("restaurants", "cafe", "market", "other", "car_services"))
        Lz[k] = [b for b in BRANDS if b in txt]
    return U, Lz


def measure_anchor(vol, new, attr, U, Lz):
    """هل عددُ المستأجرين يفسّر لترَ المحطة؟ — تُقاس الرافعةُ قبل استعمالها

       تُفحَص على الناضجةِ التي لها سجلُّ وحدات، ويُحتفَظ بالنقيض صريحاً:
       محطاتٌ بلا مستأجرٍ واحدٍ تبيع أكثرَ من محطاتٍ بسبعةَ عشرَ مستأجراً.
       فالارتباطُ قائم، والسببيّةُ غيرُ مُثبَتة.
    """
    pts = []
    for c, v in vol.items():
        if c in new or c not in attr or c not in U:
            continue
        t = totals(v, 2026)
        if not all(x > 0 for x in t[:NACT]):
            continue
        pts.append(dict(code=c, name=attr[c]["name"], reg=attr[c]["reg"],
                        city=attr[c]["city"],
                        rate=st.mean([t[i] / DAYS[i] for i in range(NACT)]),
                        leased=U[c]["leased"], built=U[c]["built"],
                        brands=len(Lz.get(c, []))))
    if len(pts) < 20:
        return None

    def corr(xs, ys):
        mx, my = st.mean(xs), st.mean(ys)
        sx = math.sqrt(sum((x - mx) ** 2 for x in xs))
        sy = math.sqrt(sum((y - my) ** 2 for y in ys))
        return sum((x - mx) * (y - my) for x, y in zip(xs, ys)) / (sx * sy) \
            if sx * sy else 0.0

    xs = [p["leased"] for p in pts]
    ys = [p["rate"] for p in pts]
    r_leased = corr(xs, ys)
    r_brands = corr([p["brands"] for p in pts], ys)
    mx, my = st.mean(xs), st.mean(ys)
    slope = sum((x - mx) * (y - my) for x, y in zip(xs, ys)) \
        / sum((x - mx) ** 2 for x in xs)
    inter = my - slope * mx
    # منحنى الفئات — أقوى من الخطِّ أمام الشواذّ، ويُعرَض للمراجعة
    BK = [(0, 1, "صفر"), (1, 4, "١–٣"), (4, 8, "٤–٧"),
          (8, 14, "٨–١٣"), (14, 999, "١٤+")]
    base = st.median(ys)
    buckets = []
    for lo, hi, lbl in BK:
        g = [p["rate"] for p in pts if lo <= p["leased"] < hi]
        if g:
            buckets.append([lbl, len(g), round(st.median(g)),
                            round(st.median(g) / base, 3)])
    tl = sum(U[p["code"]]["leased"] for p in pts)
    tn = sum(U[p["code"]]["built"] for p in pts)
    pu = collections.defaultdict(list)
    for p in pts:
        pu["منطقة " + p["reg"]].append(p["leased"])
        pu["مدينة " + p["city"].strip().title()].append(p["leased"])
    return dict(
        n=len(pts), r_leased=round(r_leased, 3), r2=round(r_leased ** 2, 3),
        r_brands=round(r_brands, 3), r2_brands=round(r_brands ** 2, 3),
        slope=round(slope), intercept=round(inter),
        net_units=st.median(xs), buckets=buckets, base=round(base),
        occ=round(tl / tn, 4) if tn else 0,
        cap_rate=round(max(ys)),
        cap_name=max(pts, key=lambda p: p["rate"])["name"],
        max_units=max(xs),
        peer_units={k: st.median(v) for k, v in pu.items() if len(v) >= 3},
        # النقيضُ — يُكتب ولا يُحذَف
        zero_high=[[p["name"], round(p["rate"])] for p in pts
                   if p["leased"] == 0 and p["rate"] > 40000],
        many_low=[[p["name"], round(p["rate"]), p["leased"], p["brands"]]
                  for p in pts if p["leased"] >= 10 and p["rate"] < 20000],
        richest=[[p["name"], p["brands"], p["leased"], round(p["rate"])]
                 for p in sorted(pts, key=lambda x: -x["brands"])[:6]])


def measure_mix(vol, new, attr):
    """حصصُ المنتجاتِ الفعليّةُ بالمنطقة — ينا–سبت ٢٠٢٦"""
    reg = collections.defaultdict(collections.Counter)
    for c, v in vol.items():
        if c in new or c not in attr:
            continue
        a = totals(v, 2026)
        if not all(x > 0 for x in a[:NACT]):
            continue
        for p in v:
            reg[attr[c]["reg"]][p] += sum(v[p][2026][:NACT])
    out, net = {}, collections.Counter()
    for rg, cnt in reg.items():
        s = sum(cnt[p] for p in PRODS)
        if s > 0:
            out[rg] = {p: round(cnt[p] / s, 4) for p in PRODS}
        for p in PRODS:
            net[p] += cnt[p]
    sn = sum(net[p] for p in PRODS)
    mnet = {p: round(net[p] / sn, 4) for p in PRODS}
    return out, mnet


def build(path):
    attr, new, vol = read(path)
    curve, nramp = measure_ramp(vol, new)
    rate, city, reg = measure_peers(vol, new, attr)
    mix, mnet = measure_mix(vol, new, attr)
    cls = measure_classes(vol, new, attr)
    prof = measure_profiles(vol, new, attr)
    U, Lz = read_units()
    anc = measure_anchor(vol, new, attr, U, Lz)
    net = st.median(list(rate.values()))
    top_mature = round(max(rate.values()))
    top_name = attr[max(rate, key=rate.get)]["name"]

    def peer(s):
        k = s["city"].strip().lower()
        if len(city.get(k, [])) >= 3:
            return st.median(city[k]), "مدينة " + k.title(), len(city[k])
        if len(reg.get(s["reg"], [])) >= 3:
            return st.median(reg[s["reg"]]), "منطقة " + s["reg"], len(reg[s["reg"]])
        return net, "الشبكة", len(rate)

    out = []
    for c, s in sorted(new.items(), key=lambda x: (x[1]["open"], x[0])):
        d = dt.date.fromisoformat(s["open"])
        ov = OVERRIDE.get(c, {})
        k = ov.get("cls")
        pr, src, n = peer(s)
        pm = mix.get(s["reg"], mnet)
        ub = U.get(c, {}).get("built", 0)
        ur = 0
        if anc and ub:
            ur = min(anc["intercept"] + anc["slope"] * round(ub * anc["occ"]),
                     anc["cap_rate"])
        basis = ov.get("basis") or ("فئةٌ مقيسة" if k in cls else "نظيرُ الجوار")
        if basis == "من عددِ الوحدات" and ur:
            r = ur
        elif k and k in cls:
            r = cls[k]["rate"]
        else:
            r = pr
        m = cls[k]["mix"] if (k and k in cls) else pm
        bu = {p: vol[c][p][2027] if p in vol[c] else [0.0] * 12 for p in PRODS}
        a26 = {p: vol[c][p][2026] if p in vol[c] else [0.0] * 12 for p in PRODS}
        tbu = [sum(bu[p][i] for p in PRODS) for i in range(12)]
        act = [x for x in tbu if x > 0]
        # الموسمُ الافتراضيّ: المنحنى الغربيُّ المقيسُ لمحطاتِ الغربية، إذ
        # قيس على ثماني محطاتٍ غربيةٍ لها سنتان فعليّتان. وما عداها «محايد»
        # لأنّه غيرُ مقيسٍ — لا لأنّ موسمَه معدوم.
        pf = ov.get("prof") or ("غربيٌّ مقيس" if s["reg"] == "Western" else "محايد")
        out.append(dict(
            code=c, name=s["name"], model=s["model"], reg=s["reg"], city=s["city"],
            open=s["open"],
            # «شهرُه»: موضعُ الافتتاح نسبةً إلى يناير ٢٠٢٧، مرقَّمٌ من صفر
            off=(d.year - 2027) * 12 + d.month - 1,
            day=d.day,                      # يومُ الافتتاح — يُحسب به كسرُ شهرِه
            rate=round(r), src=src, n=n, cls=k or "", prof=pf,
            peer_rate=round(pr), cls_rate=round(cls[k]["rate"]) if k in cls else 0,
            basis=basis,
            units_built=U.get(c, {}).get("built", 0),
            units_leased=U.get(c, {}).get("leased", 0),
            brands=len(Lz.get(c, [])),
            peer_units=(anc or {}).get("peer_units", {}).get(src, 0),
            why=ov.get("why", ""), mix=m, bu=bu, a26=a26,
            has26=any(sum(a26[p]) for p in PRODS),
            # معدّلُ الموازنةِ الضمنيّ، وهل هي رقمٌ واحدٌ مكرَّرٌ في كلِّ شهر
            bu_rate=round(max(tbu) / 31) if act else 0,
            bu_flat=len({round(x) for x in act}) == 1 and len(act) > 1))

    D = dict(stations=out, ramp=curve, days=DAYS, frac_sample=FRAC_SAMPLE,
             peer_city={k: [round(st.median(v)), len(v)]
                        for k, v in city.items() if len(v) >= 3},
             peer_reg={k: [round(st.median(v)), len(v)] for k, v in reg.items()},
             peer_net=round(net), n_mature=len(rate), n_ramp=nramp,
             mix_reg=mix, mix_net=mnet, classes=cls, profiles=prof,
             anchor=anc, top_mature=top_mature, top_name=top_name)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(D, f, ensure_ascii=False, indent=1)

    bu = sum(sum(sum(s["bu"][p]) for p in PRODS) for s in out)
    print(f"✓ {OUT}")
    print(f"  محطات {len(out)} · ناضجةٌ للمعايرة {len(rate)} · "
          f"افتتاحاتٌ قيس بها التدرّج {nramp}")
    print(f"  منحنى التدرّج: {' · '.join(f'{x:.2f}' for x in curve)}")
    print(f"  وسيطُ الشبكة {net:,.0f} ل/يوم · موازنةُ الـ٣٨ {bu / 1e6:.1f}م لتر")
    for k, v in cls.items():
        print(f"  فئة «{k}»: وسيط {v['rate']:,} ل/يوم · ن={v['n']} · "
              f"ر٣ {v['q3']:,} · أعلى {v['top']:,} · ديزل "
              f"{v['mix']['diesel'] * 100:.0f}٪")
    print(f"  منحنى «صيفيٌّ مقيس» (ن={prof['summer_n']}): "
          f"{' · '.join(f'{x:.2f}' for x in prof['summer'][:9])} "
          f"| أكت–ديس مفترضٌ {prof['summer'][9]:.2f}")
    print(f"  أعلى محطةٍ قائمة: {top_name} {top_mature:,} ل/يوم")
    print(f"  موازنةٌ مسطَّحةٌ تماماً: {sum(1 for s in out if s['bu_flat'])} محطة "
          f"من {len(out)}")
    for c, o in OVERRIDE.items():
        print(f"  OVERRIDE {c}: {o.get('cls', '')} "
              f"{'· ' + o['prof'] if 'prof' in o else ''}")
    if anc:
        print(f"  رافعةُ المستأجرين (ن={anc['n']}): وحداتٌ مؤجَّرة r="
              f"{anc['r_leased']:+.3f} R²={anc['r2']:.2f} · "
              f"علاماتٌ كبرى r={anc['r_brands']:+.3f} R²={anc['r2_brands']:.2f}")
        print(f"    لتر/يوم = {anc['intercept']:,} + {anc['slope']:,} × وحدة "
              f"· إشغالُ الناضجة {anc['occ'] * 100:.1f}٪ · سقفٌ "
              f"{anc['cap_rate']:,} ({anc['cap_name']})")
        print(f"    والنقيض: {len(anc['zero_high'])} بلا مستأجرٍ فوق ٤٠ ألف · "
              f"{len(anc['many_low'])} بعشرِ وحداتٍ دون ٢٠ ألف")
        nu = [s for s in out if s["units_built"]]
        print(f"    وللـ{len(nu)} محطةً من الـ٣٨ سجلُّ وحدات — أكبرها "
              + " · ".join(f"{x['name'][:12]} {x['units_built']}"
                           for x in sorted(nu, key=lambda y: -y["units_built"])[:4]))
    return D


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit("الاستخدام: python3.13 -I data/read_new_stations.py <ملف الموازنة>")
    build(sys.argv[1])
