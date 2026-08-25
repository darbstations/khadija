# -*- coding: utf-8 -*-
"""نموذج خطة المبيعات — التقسيم والمستهدفات وحالات المحطات
   المصدر: outlets/data/network-sales.csv (كاش إن · ٥٥ محطة · يناير–يوليو ٢٠٢٦)
   يُشغَّل من جذر المستودع:  PYTHONPATH=data python3 data/sales_plan.py

   المنهج برافعتين لا واحدة:
   (أ) متوسط حجم التعبئة — بعد عزل أثر مزيج الوقود، فثلثا تباينه مزيج لا سلوك
   (ب) المعاملات — فاقد ساعات الذروة مقارنة بشكل الشريحة
"""
import csv, json, math, statistics as stx
from collections import defaultdict

SALES = "outlets/data/network-sales.csv"
HOURS = "outlets/data/network-hours.csv"
STAFF = "outlets/data/station-staff.csv"
NET = "outlets/data/network-sales.json"
FIVE = "data/five.json"
UNITS = "outlets/data/units-registry.csv"
OUT = "data/sales-plan.json"

# ── ثوابت من تقرير يوليو ٢٠٢٦ (الشبكة المشغّلة: ٣٥٬٥٣٦ ألف ÷ ٣١٠٫٧ مليون لتر)
MARGIN = 0.1144          # هامش المساهمة — ريال لكل لتر
OPEX_NET = 28704000 / 661_500_000    # مصاريف التشغيل ÷ لترات الشبكة كلها
OPEX_OP = 28704000 / 310_700_000     # ÷ لترات المشغّلة وحدها
CLOSE = 0.40             # نغلق ٤٠٪ من الفجوة خلال ٦ أشهر
VAT = 1.15
BOX = 0.45               # كلفة علبة المناديل — مؤكَّدة من الإدارة التجارية
THRESHOLD = 50.0         # عتبة الهدية — ريال

# أسعار المضخة الفعلية من مبيعاتنا (شاملة الضريبة)
PRICE = {"بنزين ٩١": 2.191, "بنزين ٩٥": 2.340, "ديزل": 1.796}
MIX = {"بنزين ٩١": 0.555, "بنزين ٩٥": 0.360, "ديزل": 0.085}

# ── مصطلحات القطاع: نوع الموقع، وفئة الإنتاجية بمقياس MLPA
SEGS = ["الطرق السريعة", "الأساطيل التجارية", "حضرية", "الأحياء", "محدودة الإنتاجية"]
SEGDEF = {
    "الطرق السريعة":     ("Highway / Transit", "الديزل ٣٠٪ فأكثر من الحجم",
                          "شاحنات وأساطيل عابرة · سلة كبيرة"),
    "الأساطيل التجارية": ("Fleet / Commercial", "الدفع المؤسسي ١٠٪ فأكثر",
                          "عملاء متعاقدون · بطاقات وتطبيقات"),
    "حضرية":             ("Urban", "بنزين · ٩ MLPA فأكثر",
                          "أفراد المدينة · تردد عالٍ وسلة صغيرة"),
    "الأحياء":           ("Neighbourhood", "بنزين · ٤٫٥–٩ MLPA",
                          "أحياء سكنية · نمو بالتردد"),
    "محدودة الإنتاجية":  ("Low-Throughput", "دون ٤٫٥ MLPA",
                          "طلب محدود · الأولوية ضبط التكلفة"),
}
# المحرّك الأول لكل شريحة والفعل التجاري — بلا حشو
PLAY = {
    "الطرق السريعة":     ("سلة التعبئة والتعاقد", "عقود ديزل متدرجة · خدمات شاحنات · ممر سريع"),
    "الأساطيل التجارية": ("عمق الحساب المؤسسي", "مدير حساب · حدود ائتمان · تقرير استهلاك"),
    "حضرية":             ("تغطية ساعات الذروة", "طاقم الذروة · توفّر المضخات · سرعة الخدمة"),
    "الأحياء":           ("التردد", "ساعات العمل · مزيج المستأجرين · سرعة الذروة"),
    "محدودة الإنتاجية":  ("حماية الربحية", "ضبط العمالة · تأجير الوحدات · لا إنفاق تسويقي"),
}

# ── مصفوفة الأحداث: الحالة تُرصد من البيانات، ولكل حالة فعل واحد ومالك واحد
EVENTS = [
    ("فجوة مطابقة تتجاوز ٥٪", "تحقيق ورقابة — لا حملة",
     "جرد خزانات · معايرة عدادات · مطابقة إشعارات", "رقابة", "hi"),
    ("تغيير مسار أو مدخل", "هندسة وصول — لا حملة",
     "تصحيح المدخل على الخرائط · لافتة على المسار · إعادة تحجيم العمالة", "وصول", "hi"),
    ("وسيلة الدفع غير مسجّلة", "تفعيل النظام قبل أي حملة",
     "ضبط نقاط البيع · تدريب الوردية · مؤشر في حافز العامل", "بيانات", "hi"),
    ("تغطية خزان دون ٤ أيام", "تأمين الإمداد قبل رفع الطلب",
     "حد أدنى تعاقدي · ناقل بديل · دراسة خزان", "إمداد", "hi"),
    ("فاقد في ساعات الذروة", "تغطية لا حملة",
     "طاقم الذروة · فحص المضخات · تدرّج التسليم", "معاملات", "md"),
    ("تعبئة دون ما يفسّره مزيجها", "رفع السلة",
     "علبة المناديل عند العتبة · نص موحّد · لوحة عند المضخة", "سلة", "md"),
    ("حصة الديزل ٣٠٪ فأكثر", "تعاقد أساطيل",
     "حصر ناقلي المنطقة · سعر متدرج · تسجيل في التطبيق", "أساطيل", "md"),
    ("عشر وحدات شاغرة فأكثر", "باقة تأجير حسب الفئة",
     "التدريج لتحت التنفيذ · النسبة للمشغّلة · المسوّق للامتياز", "تأجير", "md"),
]

# ── باقات التأجير: منتج يُباع، لا حملة
PACKAGES = [
    ("نسبة من المبيعات", "قائمة", "المشغّلة — ١٣١ وحدة",
     "إيجار أساسي منخفض + نسبة من مبيعات المستأجر", "ربط نقاط بيع المستأجر أو إقرار مُدقَّق"),
    ("التدريج", "قائمة", "تحت التنفيذ — ٦٨٩ وحدة",
     "سنة أولى مخفّضة ترتفع تدريجياً مع نمو الحركة", "جدول افتتاحات معتمد"),
    ("المسوّق الخارجي", "قائمة", "الامتياز — ٢٦٩ وحدة",
     "عمولة على العقد المُوقَّع لا على العرض", "قائمة وحدات بأسعارها + عقد وساطة"),
    ("حد أدنى ونسبة", "مقترحة", "المشغّلة — بديل النسبة الصافية",
     "أرضية ثابتة تمنع تعطيل الوحدة + نسبة فوق حد المبيعات", "قرار تسعير"),
    ("إيجار مقابل تجهيز", "مقترحة", "تحت التنفيذ — أكبر مخزون",
     "المستأجر يجهّز ويُخصم من الإيجار على سنتين", "سقف كلفة تجهيز معتمد"),
    ("ترخيص إشغال", "مقترحة", "٥٧٢ محلاً عالقاً",
     "رسم شهري لعربة أو صرّاف أو خزانة — لا عقد إيجار", "نموذج ترخيص قابل للإلغاء"),
    ("المستأجر الرئيسي", "مقترحة", "حلي MK036 — ٩٢ وحدة صفر مؤجّر",
     "مشغّل واحد يأخذ كل الوحدات ويؤجّر من الباطن", "قرار سعر إجمالي"),
    ("الموسمي", "مقترحة", "مكة والجموم والطائف — ٢١٤ وحدة",
     "عقد شهرين بسعر مرتفع لموسم الحج والعمرة", "نموذج عقد قصير"),
]

# ── المنتجات الخمسة ومن يسيطر عليها
PRODUCTS = [
    ("وقود أفراد", "٥٠٨٫٨ مليون ريال", "٨١٪ من مبيعاتنا",
     "الدريس", "١٥ من ٤٠ موقعاً مرصوداً · ١٨٫٩٥٪ وطنياً", "ميداني — مكتمل"),
    ("وقود شركات", "١١٧٫٥ مليون ريال ديزل", "الدفع المؤسسي ٩٫١٪",
     "منصات الأساطيل", "سيارة ٧٫٧٥م · بترو ٤٫٢١م · جهاز خارجي ٢٤٫٤م", "جزئي — بلا حصص سوق"),
    ("عقارات وتأجير", "٦٤١ وحدة مؤجَّرة من ١٬٧٣٣", "بلا قيمة إيجارية في السجل",
     "غير معروف", "لا مسح منافسين ولا جدول إيجارات", "مفقود"),
    ("إكسسوارات سيارات", "لم تبدأ", "صفر",
     "غير معروف", "لا مسح", "مفقود"),
    ("مساحات إعلانية", "لم تبدأ", "صفر",
     "غير معروف", "لا مسح", "مفقود"),
]


def _f(v, d=0.0):
    try: return float(v)
    except (TypeError, ValueError): return d


def load():
    hours = {r["code"]: [int(_f(r[f"h{h:02d}"])) for h in range(24)]
             for r in csv.DictReader(open(HOURS, encoding="utf-8"))}
    rows = []
    for r in csv.DictReader(open(SALES, encoding="utf-8")):
        if r["vol_ok"] != "1":
            continue
        vol, vis, rev = _f(r["volume"]), _f(r["visits"]), _f(r["revenue"])
        d = _f(r["days"], 1)
        known = _f(r["pay_cash"]) + _f(r["pay_card"]) + _f(r["pay_fleet"])
        hs = hours.get(r["code"], [0] * 24)
        t = sum(hs) or 1
        rows.append(dict(
            code=r["code"], name=r["name"], region=r["region"], days=int(d),
            rev=rev, vol=vol, vis=int(vis),
            lpd=vol / d, vpd=vis / d, mlpa=vol / d * 365 / 1e6,
            lpv=vol / vis, inv=_f(r["avg_invoice"]),
            diesel=_f(r["diesel_vol"]) / vol, g95=_f(r["g95_vol"]) / vol,
            fleet=_f(r["pay_fleet"]) / known if known else 0.0,
            unknown=_f(r["pay_unknown"]) / rev if rev else 0.0,
            hours=[x / t for x in hs], peak=int(_f(r["peak_hour"])),
            night=_f(r["night_vis"]) / vis,
            jan=_f(r["jan_lpd"], None) or None, jul=_f(r["jul_lpd"], None) or None))
    return rows


def segment(r):
    if r["diesel"] >= 0.30: return SEGS[0]
    if r["fleet"] >= 0.10:  return SEGS[1]
    if r["mlpa"] >= 9.0:    return SEGS[2]
    if r["mlpa"] >= 4.5:    return SEGS[3]
    return SEGS[4]


def lever_fill(G):
    """رافعة (أ): الفجوة في حجم التعبئة بعد عزل أثر مزيج الوقود"""
    for g, rows in G.items():
        xs = [r["diesel"] for r in rows]; ys = [r["lpv"] for r in rows]
        n = len(rows); mx = sum(xs) / n; my = sum(ys) / n
        vx = sum((a - mx) ** 2 for a in xs)
        b = (sum((a - mx) * (c - my) for a, c in zip(xs, ys)) / vx) if vx > 1e-9 else 0.0
        a0 = my - b * mx
        sy = sum((c - my) ** 2 for c in ys)
        pred = [a0 + b * x for x in xs]
        r2 = 1 - sum((c - p) ** 2 for c, p in zip(ys, pred)) / sy if sy > 1e-9 else 0.0
        for r in rows:
            r["pred"] = a0 + b * r["diesel"]
            r["resid"] = r["lpv"] - r["pred"]
            r["r2"] = r2
        ref = sorted(rows, key=lambda r: -r["resid"])[max(0, len(rows) // 4)]["resid"]
        for r in rows:
            r["bench_resid"] = ref
            r["gap_fill"] = max(ref - r["resid"], 0)
            r["tgt_lpv"] = r["lpv"] + r["gap_fill"] * CLOSE
            r["upl_fill"] = r["gap_fill"] * CLOSE * r["vpd"]
            r["sar_fill"] = r["upl_fill"] * 365 * MARGIN


def lever_txn(G):
    """رافعة (ب): فاقد المعاملات في ساعات ذروة المحطة مقارنة بشكل شريحتها
       يُحتسب في ساعات الذروة وحدها — النقص في ساعات الهدوء اختلاف طلب لا فاقد"""
    for g, rows in G.items():
        ref = [stx.median([r["hours"][h] for r in rows]) for h in range(24)]
        for r in rows:
            top = set(sorted(range(24), key=lambda h: -r["hours"][h])[:8])
            peak = quiet = 0.0
            for h in range(24):
                d = r["hours"][h] - ref[h]
                if d >= 0: continue
                if h in top: peak += -d
                else: quiet += -d
            r["ref_hours"] = ref
            r["gap_peak"] = peak * r["vpd"]        # معاملة/يوم — قابل للدفاع
            r["gap_quiet"] = quiet * r["vpd"]      # معاملة/يوم — يحتاج إثباتاً
            r["sar_txn"] = r["gap_peak"] * r["lpv"] * 365 * MARGIN
            r["sar_total"] = r["sar_fill"] + r["sar_txn"]


def conditions(rows, recon, units, access):
    uv = {u["code"]: int(_f(u["total_vacant"])) for u in units}
    for r in rows:
        c = []
        g = recon.get(r["code"])
        if g is not None and abs(g) > 5: c.append("فجوة مطابقة")
        if r["code"] in access: c.append("تغيير مسار")
        if r["unknown"] > 0.98: c.append("بيانات دفع مفقودة")
        if r["gap_peak"] > 20: c.append("فاقد ذروة")
        if r["gap_fill"] > 0.5: c.append("سلة دون المتوقَّع")
        if r["diesel"] >= 0.30: c.append("ديزل مرتفع")
        if uv.get(r["code"], 0) >= 10: c.append("وحدات شاغرة")
        r["conds"] = c


def litre_economics():
    """تفكيك اللتر الواحد وتعبئة الخمسين ريالاً"""
    out = []
    for k, pr in PRICE.items():
        L = THRESHOLD / pr
        out.append(dict(fuel=k, price=pr, net_price=pr / VAT, litres=L,
                        margin=L * MARGIN, opex=L * OPEX_NET,
                        net=L * (MARGIN - OPEX_NET), box=BOX,
                        after=L * (MARGIN - OPEX_NET) - BOX, w=MIX[k]))
    wl = sum(o["litres"] * o["w"] for o in out) / sum(o["w"] for o in out)
    wn = sum(o["net"] * o["w"] for o in out) / sum(o["w"] for o in out)
    return dict(rows=out, cpl_margin=MARGIN * 100, cpl_opex=OPEX_NET * 100,
                cpl_opex_op=OPEX_OP * 100, cpl_net=(MARGIN - OPEX_NET) * 100,
                box=BOX, threshold=THRESHOLD, w_litres=wl, w_net=wn,
                breakeven_extra=BOX / (MARGIN - OPEX_NET),
                breakeven_sar=BOX / ((MARGIN - OPEX_NET) / 2.200))


def build():
    rows = load()
    G = defaultdict(list)
    for r in rows:
        r["seg"] = segment(r); G[r["seg"]].append(r)
    lever_fill(G); lever_txn(G)
    units = list(csv.DictReader(open(UNITS, encoding="utf-8")))
    recon = {"MK008": 117.6, "RY024": 36.8, "MK003": 3.1, "MK017": -3.2}
    access = {"MK002": "٧–٨ فبراير", "MK019": "٤–٥ فبراير"}
    conditions(rows, recon, units, access)
    net = json.load(open(NET, encoding="utf-8"))
    five = json.load(open(FIVE, encoding="utf-8"))
    staff = {r["code"]: (int(r["day_workers"]), int(r["eve_workers"]))
             for r in csv.DictReader(open(STAFF, encoding="utf-8"))}

    segsum = []
    for s in SEGS:
        g = G[s]
        vol = sum(x["vol"] for x in g); vis = sum(x["vis"] for x in g)
        segsum.append(dict(
            seg=s, en=SEGDEF[s][0], rule=SEGDEF[s][1], who=SEGDEF[s][2],
            n=len(g), lpd=sum(x["lpd"] for x in g), mlpa=sum(x["mlpa"] for x in g),
            vpd=sum(x["vpd"] for x in g), lpv=vol / vis,
            diesel=sum(x["diesel"] * x["vol"] for x in g) / vol,
            fleet=sum(x["fleet"] * x["rev"] for x in g) / sum(x["rev"] for x in g),
            inv=sum(x["rev"] for x in g) / vis, r2=g[0]["r2"],
            below=sum(1 for x in g if x["gap_fill"] > 0.5),
            upl_fill=sum(x["upl_fill"] for x in g), gap_peak=sum(x["gap_peak"] for x in g),
            sar_fill=sum(x["sar_fill"] for x in g), sar_txn=sum(x["sar_txn"] for x in g),
            net_txn=(vol / vis) * (MARGIN - OPEX_NET),
            driver=PLAY[s][0], action=PLAY[s][1]))

    shifts = []
    for c, (dw, ew) in staff.items():
        r = next(x for x in rows if x["code"] == c)
        top = set(sorted(range(24), key=lambda h: -r["hours"][h])[:8])
        pd_ = sum(-(r["hours"][h] - r["ref_hours"][h]) for h in top
                  if 6 <= h < 18 and r["hours"][h] < r["ref_hours"][h]) * r["vpd"]
        pe_ = sum(-(r["hours"][h] - r["ref_hours"][h]) for h in top
                  if (h >= 18 or h < 6) and r["hours"][h] < r["ref_hours"][h]) * r["vpd"]
        day = sum(r["hours"][h] for h in range(6, 18))
        shifts.append(dict(code=c, name=r["name"], vpd=r["vpd"], dw=dw, ew=ew,
                           day=r["vpd"] * day, eve=r["vpd"] * (1 - day),
                           load_d=r["vpd"] * day / dw, load_e=r["vpd"] * (1 - day) / ew,
                           tgt_d=pd_, tgt_e=pe_, peak=r["peak"]))
    shifts.sort(key=lambda x: -x["vpd"])

    T = dict(n=len(rows), mlpa=sum(r["mlpa"] for r in rows), lpd=sum(r["lpd"] for r in rows),
             upl_fill=sum(r["upl_fill"] for r in rows),
             gap_peak=sum(r["gap_peak"] for r in rows),
             gap_quiet=sum(r["gap_quiet"] for r in rows),
             sar_fill=sum(r["sar_fill"] for r in rows),
             sar_txn=sum(r["sar_txn"] for r in rows),
             below=sum(1 for r in rows if r["gap_fill"] > 0.5))
    T["sar"] = T["sar_fill"] + T["sar_txn"]

    for r in rows:
        r.pop("hours", None); r.pop("ref_hours", None)
    D = dict(
        margin=MARGIN, opex=OPEX_NET, close=CLOSE, box=BOX, threshold=THRESHOLD,
        network=dict(stations=net["stations"], revenue=net["revenue"], volume=net["volume"],
                     visits=net["visits"], inv=net["avg_invoice"],
                     days=sum(r["days"] for r in rows), fuels=net["fuels"], pays=net["pays"],
                     monthly=net["monthly"], hours=net["hours"],
                     raw_rows=net["raw_rows"], files=net["files"]),
        segments=segsum, stations=sorted(rows, key=lambda x: -x["sar_total"]),
        shifts=shifts, events=EVENTS, packages=PACKAGES, products=PRODUCTS,
        litre=litre_economics(), access=access,
        five=[dict(code=f["code"], name=f["name"], rating=f["rating"], n=f["nComp"],
                   avg=f["compAvg"], near=f["nearDist"], who=f["nearName"],
                   density=f["density"]) for f in five],
        units=dict(n=len(units),
                   total=sum(int(_f(u["total_n"])) for u in units),
                   leased=sum(int(_f(u["total_leased"])) for u in units),
                   vacant=sum(int(_f(u["total_vacant"])) for u in units),
                   shops=sum(int(_f(u["shop_n"])) for u in units),
                   shops_leased=sum(int(_f(u["shop_leased"])) for u in units),
                   cats=[dict(cat=c,
                              n=sum(1 for u in units if u["category"] == c),
                              vacant=sum(int(_f(u["total_vacant"]))
                                         for u in units if u["category"] == c))
                         for c in ("مشغّلة", "تحت التنفيذ", "امتياز")],
                   top=sorted([dict(code=u["code"], name=u["name"], cat=u["category"],
                                    n=int(_f(u["total_n"])), leased=int(_f(u["total_leased"])),
                                    vacant=int(_f(u["total_vacant"]))) for u in units],
                              key=lambda x: -x["vacant"])[:6]),
        totals=T)
    json.dump(D, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    return D


if __name__ == "__main__":
    D = build()
    t = D["totals"]
    print(f"محطات: {t['n']} · {t['mlpa']:.0f} MLPA · دون المتوقَّع: {t['below']}")
    print(f"رافعة التعبئة : {t['upl_fill']:>8,.0f} لتر/يوم = {t['sar_fill']:>10,.0f} ر/سنة")
    print(f"رافعة المعاملات: {t['gap_peak']:>8,.0f} معاملة/يوم = {t['sar_txn']:>10,.0f} ر/سنة")
    print(f"الإجمالي المؤكَّد: {t['sar']:,.0f} ريال/سنة "
          f"(وفاقد ساعات الهدوء {t['gap_quiet']:,.0f} معاملة/يوم غير مؤكَّد)")
    for s in D["segments"]:
        print(f"  {s['seg']:<18}{s['n']:>3} محطة · {s['mlpa']:>6.1f} MLPA · R²={s['r2']:.2f}"
              f" · تعبئة {s['sar_fill']:>9,.0f} · معاملات {s['sar_txn']:>9,.0f}")
    L = D["litre"]
    print(f"\nاللتر: هامش {L['cpl_margin']:.2f} هللة − تشغيل {L['cpl_opex']:.2f} "
          f"= صافي {L['cpl_net']:.2f} · تعادل العلبة {L['breakeven_extra']:.1f} لتر "
          f"({L['breakeven_sar']:.1f} ريال إنفاق إضافي)")
    print("كُتب", OUT)
