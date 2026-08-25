# -*- coding: utf-8 -*-
"""نموذج خطة المبيعات — التقسيم والمستهدفات وحالات المحطات
   المصدر: outlets/data/network-sales.csv (كاش إن ، ٥٥ محطة · يناير–يوليو ٢٠٢٦)
   يُشغَّل من جذر المستودع:  PYTHONPATH=data python3 data/sales_plan.py
"""
import csv, json, statistics as stx
from collections import defaultdict

SALES = "outlets/data/network-sales.csv"
NET = "outlets/data/network-sales.json"
FIVE = "data/five.json"
UNITS = "outlets/data/units-registry.csv"
OUT = "data/sales-plan.json"

# هامش المساهمة للشبكة المشغّلة — تقرير يوليو ٢٠٢٦ (٣٥٬٥٣٦ ألف ÷ ٣١٠٫٧ مليون لتر)
MARGIN = 0.1131          # ريال لكل لتر
CLOSE = 0.40             # نغلق ٤٠٪ من فجوة المعيار خلال ٦ أشهر
VAT = 1.15

SEGS = ["محاور الطريق", "مراكز الأساطيل", "حضرية كبرى", "حضرية متوسطة", "محلية صغيرة"]

SEGDEF = {
    "محاور الطريق":   ("الديزل ٣٠٪ فأكثر من الحجم", "شاحنات وأساطيل عابرة · سلة كبيرة · ذروة نهارية"),
    "مراكز الأساطيل": ("الدفع المؤسسي ١٠٪ فأكثر", "عملاء متعاقدون · بطاقات وتطبيقات · تكرار عالٍ"),
    "حضرية كبرى":     ("بنزين، ٢٥ ألف لتر/يوم فأكثر", "أفراد المدينة · تردد عالٍ وسلة صغيرة"),
    "حضرية متوسطة":   ("بنزين، من ١٢ إلى ٢٥ ألف لتر/يوم", "أحياء سكنية · نمو بالتردد لا بالسلة"),
    "محلية صغيرة":    ("أقل من ١٢ ألف لتر/يوم", "طلب محدود · الأولوية ضبط التكلفة"),
}

# خطة كل شريحة: (المحرّك، الفعل التجاري، الحملة الأنسب)
PLAY = {
    "محاور الطريق": ("رفع سلة التعبئة والتعاقد",
                     "عقود ديزل بأسعار متدرجة · خدمات الشاحنات (إطارات وزيوت واستراحة) · ممر تعبئة سريع",
                     "علبة مناديل عند ٥٠ لتراً فأكثر · كوبون استراحة"),
    "مراكز الأساطيل": ("تعميق الحساب المؤسسي",
                       "مدير حساب لكل عميل · حدود ائتمان · تقرير استهلاك شهري للعميل",
                       "عرض تسجيل أسطول: خصم خدمات للمركبة"),
    "حضرية كبرى": ("رفع اللتر لكل زيارة",
                   "توجيه الممتاز ٩٥ · تعبئة كاملة بدل جزئية · ربط المتجر والمغسلة",
                   "علبة مناديل عند ٣٠ لتراً فأكثر · باقة غسيل"),
    "حضرية متوسطة": ("رفع التردد",
                     "حملة حي · شراكات مع مستأجري المحطة · ساعات ذروة مدعّمة",
                     "بطاقة الحي: الزيارة الخامسة غسيل مجاني"),
    "محلية صغيرة": ("حماية الربحية",
                    "ضبط العمالة على الطلب · تأجير الوحدات الشاغرة · لا إنفاق تسويقي",
                    "فعالية مجتمعية واحدة كل ربع بتمويل مورّد"),
}

# مصفوفة الأحداث — الحالة تُكتشف من البيانات ثم تُقابَل بفعل واحد محدد
EVENTS = [
    ("فجوة مطابقة تتجاوز ٥٪", "تحقيق ورقابة — لا حملة",
     "جرد خزانات فوري · معايرة عدادات · مطابقة إشعارات الاستلام", "رقابة", "hi"),
    ("انهيار حجم يتجاوز ٢٠٪ في ٧ أشهر", "حملة استرجاع موجّهة",
     "مسح المنافس الأقرب · عرض سعري/خدمي ٦ أسابيع · قياس أسبوعي", "استرجاع", "md"),
    ("وسيلة الدفع غير مسجّلة", "تفعيل النظام قبل أي حملة",
     "ضبط إعداد نقاط البيع · تدريب الوردية · مؤشر تسجيل في حافز العامل", "بيانات", "hi"),
    ("لتر/زيارة دون معيار الشريحة", "حملة رفع السلة",
     "علبة مناديل عند العتبة · نص موحّد للعامل · لوحة عتبة عند المضخة", "سلة", "md"),
    ("حصة الديزل ٣٠٪ فأكثر", "تعاقد أساطيل",
     "حصر ناقلي المنطقة · عرض تعاقد بسعر متدرج · تسجيل في التطبيق", "أساطيل", "md"),
    ("منافس دون ٥٠٠ متر", "تفوق خدمي معلن",
     "سرعة تعبئة مضمونة · نظافة دورات المياه · إبراز التقييم على الخرائط", "تنافس", "md"),
    ("عشر وحدات شاغرة فأكثر", "يوم مفتوح للمستأجرين",
     "عرض مساحة مقابل نسبة · وسيط محلي · باقة افتتاح", "تأجير", "md"),
    ("تغطية خزان دون ٤ أيام", "تأمين الإمداد قبل رفع الطلب",
     "حد أدنى تعاقدي للتحميلات · ناقل بديل مؤهّل · دراسة خزان إضافي", "إمداد", "hi"),
    ("محطة تحت التنفيذ", "تأجير مسبق قبل الافتتاح",
     "تسويق الوحدات من الشهر −٣ · باقة مستأجر مؤسِّس", "تأجير", "md"),
]


def _f(v, d=0.0):
    try: return float(v)
    except (TypeError, ValueError): return d


def load():
    rows = []
    for r in csv.DictReader(open(SALES, encoding="utf-8")):
        if r["vol_ok"] != "1":
            continue
        vol, vis, rev = _f(r["volume"]), _f(r["visits"]), _f(r["revenue"])
        d = _f(r["days"], 1)
        known = _f(r["pay_cash"]) + _f(r["pay_card"]) + _f(r["pay_fleet"])
        rows.append(dict(
            code=r["code"], name=r["name"], region=r["region"], days=int(d),
            rev=rev, vol=vol, vis=int(vis),
            lpd=vol / d, vpd=vis / d, rpd=rev / d,
            lpv=vol / vis, inv=_f(r["avg_invoice"]),
            diesel=_f(r["diesel_vol"]) / vol, g95=_f(r["g95_vol"]) / vol,
            fleet=_f(r["pay_fleet"]) / known if known else 0.0,
            unknown=_f(r["pay_unknown"]) / rev if rev else 0.0,
            night=_f(r["night_vis"]) / vis, peak=int(_f(r["peak_hour"])),
            jan=_f(r["jan_lpd"], None) or None, jul=_f(r["jul_lpd"], None) or None))
    return rows


def segment(r):
    if r["diesel"] >= 0.30: return SEGS[0]
    if r["fleet"] >= 0.10:  return SEGS[1]
    if r["lpd"] >= 25000:   return SEGS[2]
    if r["lpd"] >= 12000:   return SEGS[3]
    return SEGS[4]


def plan(rows):
    """يضع لكل محطة معيار شريحتها ومستهدفها والقيمة السنوية للفجوة"""
    G = defaultdict(list)
    for r in rows:
        r["seg"] = segment(r); G[r["seg"]].append(r)
    for s, g in G.items():
        bench = stx.quantiles([x["lpv"] for x in g], n=4)[2]   # الربيع الأعلى داخل الشريحة
        for r in g:
            r["bench"] = bench
            r["gap"] = max(bench - r["lpv"], 0)
            r["tgt_lpv"] = r["lpv"] + r["gap"] * CLOSE
            r["upl_lpd"] = r["gap"] * CLOSE * r["vpd"]
            r["upl_sar"] = r["upl_lpd"] * 365 * MARGIN
    return G


def conditions(rows, recon, units):
    """يوسم كل محطة بالحالات المنطبقة عليها — هي التي تختار الفعل"""
    uv = {u["code"]: int(_f(u["total_vacant"])) for u in units}
    for r in rows:
        c = []
        g = recon.get(r["code"])
        if g is not None and abs(g) > 5: c.append("فجوة مطابقة")
        if r["jan"] and r["jul"] and r["jul"] / r["jan"] - 1 < -0.20: c.append("انهيار حجم")
        if r["unknown"] > 0.98: c.append("بيانات دفع مفقودة")
        if r["gap"] > 0: c.append("سلة دون المعيار")
        if r["diesel"] >= 0.30: c.append("ديزل مرتفع")
        if uv.get(r["code"], 0) >= 10: c.append("وحدات شاغرة")
        r["conds"] = c
    return rows


def worker_targets(rows):
    """مستهدف العامل: زيارات الوردية من منحنى الطلب، والسلة من معيار الشريحة"""
    out = []
    for r in rows:
        eve = r["night"]                       # حصة ١٨:٠٠–٠٦:٠٠ من الزيارات
        out.append(dict(code=r["code"], name=r["name"], seg=r["seg"],
                        vpd=r["vpd"], day_vis=r["vpd"] * (1 - eve), eve_vis=r["vpd"] * eve,
                        eve_share=eve, lpv=r["lpv"], tgt_lpv=r["tgt_lpv"], peak=r["peak"]))
    return out


def build():
    rows = load()
    G = plan(rows)
    units = list(csv.DictReader(open(UNITS, encoding="utf-8")))
    recon = {"MK008": 117.6, "RY024": 36.8, "MK003": 3.1, "MK017": -3.2}
    conditions(rows, recon, units)
    net = json.load(open(NET, encoding="utf-8"))
    five = json.load(open(FIVE, encoding="utf-8"))

    segsum = []
    for s in SEGS:
        g = G[s]
        vol = sum(x["vol"] for x in g); vis = sum(x["vis"] for x in g)
        segsum.append(dict(
            seg=s, n=len(g), lpd=sum(x["lpd"] for x in g), vpd=sum(x["vpd"] for x in g),
            lpv=vol / vis, diesel=sum(x["diesel"] * x["vol"] for x in g) / vol,
            fleet=sum(x["fleet"] * x["rev"] for x in g) / sum(x["rev"] for x in g),
            night=sum(x["night"] * x["vis"] for x in g) / vis,
            inv=sum(x["rev"] for x in g) / vis,
            bench=g[0]["bench"], below=sum(1 for x in g if x["gap"] > 0),
            upl=sum(x["upl_lpd"] for x in g), sar=sum(x["upl_sar"] for x in g),
            rule=SEGDEF[s][0], who=SEGDEF[s][1],
            driver=PLAY[s][0], action=PLAY[s][1], camp=PLAY[s][2]))

    D = dict(
        margin=MARGIN, close=CLOSE, vat=VAT,
        network=dict(stations=net["stations"], revenue=net["revenue"], volume=net["volume"],
                     visits=net["visits"], inv=net["avg_invoice"],
                     days=sum(r["days"] for r in rows), fuels=net["fuels"],
                     pays=net["pays"], monthly=net["monthly"], hours=net["hours"],
                     raw_rows=net["raw_rows"], files=net["files"]),
        segments=segsum, stations=sorted(rows, key=lambda x: -x["upl_lpd"]),
        workers=worker_targets(rows), events=EVENTS, five=[
            dict(code=f["code"], name=f["name"], rating=f["rating"], n=f["nComp"],
                 avg=f["compAvg"], near=f["nearDist"], who=f["nearName"],
                 density=f["density"]) for f in five],
        units=dict(
            n=len(units),
            total=sum(int(_f(u["total_n"])) for u in units),
            leased=sum(int(_f(u["total_leased"])) for u in units),
            vacant=sum(int(_f(u["total_vacant"])) for u in units),
            shops=sum(int(_f(u["shop_n"])) for u in units),
            shops_leased=sum(int(_f(u["shop_leased"])) for u in units),
            top=sorted([dict(code=u["code"], name=u["name"], cat=u["category"],
                             n=int(_f(u["total_n"])), leased=int(_f(u["total_leased"])),
                             vacant=int(_f(u["total_vacant"])))
                        for u in units], key=lambda x: -x["vacant"])[:6]),
        totals=dict(lpd=sum(r["lpd"] for r in rows), upl=sum(r["upl_lpd"] for r in rows),
                    sar=sum(r["upl_sar"] for r in rows),
                    below=sum(1 for r in rows if r["gap"] > 0), n=len(rows)))
    json.dump(D, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    return D


if __name__ == "__main__":
    D = build()
    t = D["totals"]
    print(f"محطات في النموذج: {t['n']} · دون المعيار: {t['below']}")
    print(f"القاعدة {t['lpd']:,.0f} لتر/يوم · الرفع {t['upl']:,.0f} ({t['upl']/t['lpd']*100:.1f}٪)"
          f" · القيمة {t['sar']:,.0f} ريال/سنة")
    for s in D["segments"]:
        print(f"  {s['seg']:<16}{s['n']:>3} محطة · {s['lpd']:>9,.0f} ل/يوم · "
              f"معيار {s['bench']:>5.1f} · دونه {s['below']:>2} · رفع {s['upl']:>7,.0f}")
    print("كُتب", OUT)
