# -*- coding: utf-8 -*-
"""تحليل المنافسين لكل منتج من منتجاتنا الخمسة

   لكل منتج منافس مختلف، ومقياس مختلف، ومصدر بيانات مختلف:
   ① وقود الأفراد   — محطات مجاورة   · مسح ميداني من خرائط جوجل (يوليو ٢٠٢٦)
   ② وقود الشركات   — منصّات الأساطيل · وسائل الدفع في نقاط البيع
   ③ العقار والتأجير — بديل المستأجر   · سجل الوحدات ١٨٦ محطة
   ④ الإكسسوارات    — مستأجرونا أنفسهم · سجل الوحدات
   ⑤ المساحات       — مشغّلو الإعلان   · حجم الجمهور من المعاملات

   يُشغَّل من جذر المستودع:  PYTHONPATH=data python3 data/competitors.py
"""
import csv, json, re, math, collections, statistics as stx

FIVE = "data/five.json"
UNITS = "outlets/data/units-registry.csv"
SALES = "outlets/data/network-sales.csv"
COMPLAINTS = "outlets/data/complaints.csv"
DASH = "data/network-payments.json"
OUT = "data/competitors.json"
PLAN = "data/sales-plan.json"              # مخرجات نموذج الخطة
DIESEL_PRICE = 1.796      # سعر مضخة الديزل — لاشتقاق اللترات حين تكون الكمية ناقصة

# من يخدمه كل تصنيف — بلغة الشركة نفسها
SGWHO = {"حيوية": "أفراد المدينة · تردد عالٍ وسلة صغيرة",
         "خط سفر": "شاحنات ومسافرون · تعبئة كاملة وسلة كبيرة",
         "حي": "سكّان الحي · النمو بالتردد لا بالسلة",
         "مختلط": "سكّان وعابرون — طلبان في موقع واحد",
         "نائية": "طلب محدود · الأولوية ضبط التكلفة"}

# تصنيف بالدليل لا بالاسم: ما فاتورته بحجم النقدي ليس أسطولاً مهما سُمّي.
#   سيارة ١٤٩ وبترو ١٧٢ ريالاً — ثلاثة أضعاف فاتورة الشبكة، فهي تعبئة أسطول.
#   «جهاز خارجي» ٥٤٫٥١ ريالاً — بحجم النقدي ٥١٫٩٢، فهو جهاز نقاط بيع لا قناة أسطول.
PLATFORMS = ("Syarah App", "Petro App")
POS_DEVICE = ("External Device", "Smart Card")
OURS = ("CASHIN Wallet",)
PAY_AR = {"Cash": "نقدي", "Card": "شبكة (مدى)", "Petro App": "تطبيق بترو",
          "Syarah App": "تطبيق سيارة", "External Device": "جهاز خارجي",
          "Smart Card": "بطاقة ذكية", "CASHIN Wallet": "محفظة كاش إن",
          "Not Specified": "غير مسجَّلة"}

BRANDS = [("aldrees", "الدريس"), ("الدريس", "الدريس"), ("naft", "نفط"), ("نفط", "نفط"),
          ("sasco", "ساسكو"), ("ساسكو", "ساسكو"), ("petrogen", "بتروجين"),
          ("بتروجين", "بتروجين"), ("omco", "أومكو"), ("اومكو", "أومكو"),
          ("adnoc", "أدنوك"), ("توتال", "توتال")]

UNIT_TYPES = [("kiosk", "أكشاك", "عربة أو كشك في ساحة تجارية"),
              ("shop", "محلات", "محل على شارع الحي — أوفر إيجاراً وأقرب للسكن"),
              ("drive", "سيّاقة", "لا بديل: نافذة السيّاقة تحتاج مسار سيارات"),
              ("carwash", "مغاسل", "مغسلة حي مستقلة — لكن بلا حركة مرور مضمونة"),
              ("market", "أسواق", "بقالة الحي"),
              ("other", "أخرى", "—")]


def _i(r, k):
    try: return int(float(r.get(k) or 0))
    except (TypeError, ValueError): return 0


def brand(n):
    for k, v in BRANDS:
        if k in n.lower():
            return v
    return "مستقلة/أخرى"


def _m(s):
    return int(re.sub(r"[^\d]", "", str(s)) or 0)


# ═══════════════════════════════════════════════════════════ ① وقود الأفراد
def retail():
    """المنافسة الميدانية: من حولنا، وهل يفسّر وجودُه نتيجتَنا"""
    F = json.load(open(FIVE, encoding="utf-8"))
    cnt = collections.Counter()
    dist, rate, revw = collections.defaultdict(list), collections.defaultdict(list), \
        collections.defaultdict(list)
    allr = []
    for s in F:
        for c in s["competitors"]:
            b = brand(c["name"]); cnt[b] += 1
            dist[b].append(_m(c["dist"]))
            try:
                r = float(c["rating"]); rate[b].append(r); allr.append(r)
            except (TypeError, ValueError):
                pass
            revw[b].append(_m(c.get("reviews", 0)))
    tot = sum(cnt.values())
    brands = [dict(brand=b, n=n, share=n / tot,
                   near=min(dist[b]), med=int(stx.median(dist[b])),
                   rating=stx.mean(rate[b]) if rate[b] else None,
                   reviews=int(stx.median(revw[b])) if revw[b] else 0)
              for b, n in cnt.most_common()]

    sites = []
    for s in F:
        rv = 0
        for line in s["swot"]["القوة"] + s["swot"]["الفرص"]:
            m = re.search(r"([\d,]+)\s*مراجعة", line)
            if m:
                rv = int(m.group(1).replace(",", "")); break
        sites.append(dict(code=s["code"], name=s["name"], rating=s["rating"],
                          reviews=rv, comp_avg=s["compAvg"], n=s["nComp"],
                          near=_m(s["nearDist"]), near_who=brand(s["nearName"]),
                          growth=float(s["growth"]), gap=s["rating"] - s["compAvg"]))

    # الحجّة: ميزتنا ثابتة على المواقع الخمسة والنتيجة تتباين — فالثابت لا يفسّر المتغيّر
    g = [x["growth"] for x in sites]
    gaps = [x["gap"] for x in sites]

    def corr(a, b):
        n = len(a); ma, mb = sum(a) / n, sum(b) / n
        va = sum((x - ma) ** 2 for x in a); vb = sum((y - mb) ** 2 for y in b)
        return sum((x - ma) * (y - mb) for x, y in zip(a, b)) / math.sqrt(va * vb)

    # شكاوانا نفسها: أين تُكسب المنافسة فعلاً حين يكون السعر منظَّماً
    C = list(csv.DictReader(open(COMPLAINTS, encoding="utf-8-sig")))
    svc = collections.Counter()
    SVC = {"سلوك العمال": "سلوك", "بطء الخدمة": "سرعة", "بطء الخدمة.": "سرعة",
           "رفض التعبئة": "رفض خدمة", "تعبئة وقود خاطئ": "خطأ تعبئة",
           "تدخين العمالة": "سلوك", "سياسة العامل مسؤول عن مضخته": "سلوك",
           "التأكد من صحة التعبئة": "ثقة", "التحقق من التعبئة": "ثقة",
           "سحب مبلغ زائد": "ثقة", "نظافة السجاد": "نظافة"}
    for r in C:
        svc[SVC.get(r["category"].strip(), "أخرى")] += 1
    n_svc = sum(v for k, v in svc.items() if k != "أخرى")

    return dict(
        brands=brands, sites=sites, n_sample=tot,
        n_total=sum(s["nComp"] for s in F),
        rate_mean=stx.mean(allr), rate_below4=sum(1 for x in allr if x < 4.0) / len(allr),
        ours_mean=stx.mean([s["rating"] for s in F]),
        gap_mean=stx.mean([s["rating"] for s in F]) - stx.mean(allr),
        growth_lo=min(g), growth_hi=max(g),
        r_gap=corr(gaps, g), r_n=corr([x["n"] for x in sites], g),
        r_dist=corr([x["near"] for x in sites], g),
        complaints=sorted(svc.items(), key=lambda x: -x[1]),
        n_complaints=len(C), svc_share=n_svc / len(C))


# ═══════════════════════════════════════════════════════ ② وقود الشركات
def corporate():
    """المنافس منصّة تملك حساب الأسطول — لا محطة على الطريق"""
    D = json.load(open(DASH, encoding="utf-8"))
    net = D["network"]
    tot = sum(p["rev"] for p in net.values())
    known = tot - net.get("Not Specified", {}).get("rev", 0)
    cash_inv = net["Cash"]["rev"] / net["Cash"]["vis"]
    pays = [dict(pay=k, ar=PAY_AR.get(k, k), rev=v["rev"], vis=v["vis"],
                 inv=v["rev"] / v["vis"] if v["vis"] else 0,
                 share=v["rev"] / tot, ours=k in OURS, plat=k in PLATFORMS,
                 pos=k in POS_DEVICE,
                 # هل السلة سلة أسطول؟ ضِعف فاتورة النقدي فأكثر
                 fleet_basket=(v["rev"] / v["vis"] if v["vis"] else 0) >= 2 * cash_inv)
            for k, v in sorted(net.items(), key=lambda x: -x[1]["rev"])]
    plat_rev = sum(p["rev"] for p in pays if p["plat"])
    plat_vis = sum(p["vis"] for p in pays if p["plat"])
    pos_rev = sum(p["rev"] for p in pays if p["pos"])
    pos_vis = sum(p["vis"] for p in pays if p["pos"])
    our_rev = sum(p["rev"] for p in pays if p["ours"])
    our_vis = sum(p["vis"] for p in pays if p["ours"])

    rows = []
    for s in D["stations"]:
        pk = {p["pay"]: p for p in s["pays"]}
        kn = sum(v["rev"] for k, v in pk.items() if k != "Not Specified")
        pl = sum(v["rev"] for k, v in pk.items() if k in PLATFORMS)
        dr = s.get("diesel_rev", 0.0)
        # ثلاث محطات كمياتها ناقصة في التصدير — تُشتقّ لتراتها من الإيراد بسعر الديزل
        dv = dr / DIESEL_PRICE if s.get("vol_partial") else s.get("diesel_vol", 0.0)
        rows.append(dict(code=s["code"], name=s["name"], rev=s["revenue"],
                         known=kn, plat=pl, share=pl / kn if kn else 0,
                         partial=bool(s.get("vol_partial")),
                         diesel=dr, diesel_vol=dv,
                         syarah=pk.get("Syarah App", {}).get("rev", 0.0),
                         petro=pk.get("Petro App", {}).get("rev", 0.0),
                         ext=pk.get("External Device", {}).get("rev", 0.0)))
    rows.sort(key=lambda r: -r["plat"])
    # قائمة الصيد: ديزل ثقيل بلا أي علاقة رقمية
    big = [r for r in rows if r["diesel"] >= 3e6]
    gap = sorted([r for r in big if r["share"] < 0.02], key=lambda r: -r["diesel"])
    return dict(pays=pays, total=tot, known=known, cash_inv=cash_inv,
                unknown_share=net.get("Not Specified", {}).get("rev", 0) / tot,
                plat_rev=plat_rev, plat_vis=plat_vis,
                plat_share=plat_rev / known, plat_inv=plat_rev / plat_vis,
                pos_rev=pos_rev, pos_vis=pos_vis, pos_inv=pos_rev / pos_vis,
                our_rev=our_rev, our_vis=our_vis,
                stations=rows[:10], big=len(big), gap=gap,
                gap_diesel=sum(r["diesel"] for r in gap),
                gap_vol=sum(r["diesel_vol"] for r in gap))


# ═════════════════════════════════════════════════════ ③④⑤ العقار وما يليه
def property_():
    """بديل المستأجر يختلف باختلاف الوحدة — والإشغال يقيس ذلك مباشرة"""
    R = list(csv.DictReader(open(UNITS, encoding="utf-8")))
    types = []
    for k, ar, rival in UNIT_TYPES:
        n = sum(_i(r, k + "_n") for r in R)
        if not n:
            continue
        l = sum(_i(r, k + "_leased") for r in R)
        types.append(dict(key=k, ar=ar, rival=rival, n=n, leased=l,
                          held=sum(_i(r, k + "_held") for r in R),
                          vacant=sum(_i(r, k + "_vacant") for r in R), occ=l / n))
    types.sort(key=lambda t: -t["occ"])
    cats = []
    for c in ("مشغّلة", "تحت التنفيذ", "امتياز"):
        g = [r for r in R if r["category"] == c]
        n = sum(_i(r, "total_n") for r in g)
        l = sum(_i(r, "total_leased") for r in g)
        cats.append(dict(cat=c, stations=len(g), n=n, leased=l,
                         vacant=sum(_i(r, "total_vacant") for r in g), occ=l / n if n else 0))
    n = sum(t["n"] for t in types); l = sum(t["leased"] for t in types)
    return dict(types=types, cats=cats, stations=len(R), n=n, leased=l,
                vacant=sum(t["vacant"] for t in types), occ=l / n,
                shops=next(t for t in types if t["key"] == "shop"),
                best=types[0])


def underwrite():
    """ورقة اكتتاب الوحدة: الإيجار يُسعَّر من دفتر المستأجر لا من متر المربّع

       الوحدة على الساحة وصولٌ إلى تيّار حركة، وقيمتها ما يحوّله المستأجر منه.
       أول مُدخَل في دفتره هو الحركة المارّة بوحدته — وهو المُدخَل الوحيد الذي
       نملكه اليوم. البقية (التحويل · السلة · الهامش) يأتي بها المستأجر أو
       مسحٌ لم يُجرَ بعد، فلا تُقدَّر هنا."""
    U = list(csv.DictReader(open(UNITS, encoding="utf-8")))
    S = {r["code"]: r for r in csv.DictReader(open(SALES, encoding="utf-8"))
         if r["vol_ok"] == "1"}
    rows = []
    for r in U:
        c, vac = r["code"], _i(r, "total_vacant")
        if not vac or c not in S:
            continue
        s = S[c]
        vpd = float(s["visits"]) / float(s["days"] or 1)
        rows.append(dict(code=c, name=r["name"], cat=r["category"], vacant=vac,
                         units=_i(r, "total_n"), vpd=vpd, per=vpd / vac,
                         shop=_i(r, "shop_vacant"), wash=_i(r, "carwash_vacant"),
                         kiosk=_i(r, "kiosk_vacant")))
    rows.sort(key=lambda x: -x["per"])
    med = stx.median([x["per"] for x in rows]) if rows else 0

    # التغطية: أي شغور نملك مُدخَله الأول أصلاً؟
    cov = []
    for cat in ("مشغّلة", "تحت التنفيذ", "امتياز"):
        g = [r for r in U if r["category"] == cat]
        v = sum(_i(r, "total_vacant") for r in g)
        vm = sum(_i(r, "total_vacant") for r in g if r["code"] in S)
        cov.append(dict(cat=cat, vacant=v, measured=vm, share=vm / v if v else 0,
                        why=("الحركة مقيسة — تُكتتب اليوم" if cat == "مشغّلة" else
                             "المحطة لم تُفتح — تُكتتب بمحطة نظيرة" if cat == "تحت التنفيذ" else
                             "المحطة تعمل ولا تصلنا معاملاتها — طلب بيانات")))
    total_v = sum(c["vacant"] for c in cov)
    return dict(rows=rows, n=len(rows), median=med,
                ready=[x for x in rows if x["per"] >= med],
                thin=[x for x in rows if x["per"] < med],
                top=rows[:6], bottom=rows[-4:],
                spread=rows[0]["per"] / rows[-1]["per"] if len(rows) > 1 else 0,
                coverage=cov, total_vacant=total_v,
                measured=sum(c["measured"] for c in cov),
                measured_share=sum(c["measured"] for c in cov) / total_v if total_v else 0,
                # المُدخَلات الأربعة لدفتر المستأجر — وأيّها عندنا
                inputs=[("الحركة المارّة بالوحدة", "زيارة/يوم لكل محطة", True),
                        ("ساعات الذروة", "منحنى ٢٤ ساعة لكل محطة", True),
                        ("تحويل الشكل من التيّار", "لكل صيغة — مطعم · مقهى · مغسلة", False),
                        ("سلة المستأجر وهامشه", "مبيعات المستأجرين شهرياً", False)])


def newcomers(P, visits_month):
    """④ الإكسسوارات و⑤ المساحات: منتجان لم يبدآ — فالسؤال من يشغل مكانهما اليوم

       لا يُدَّعى مسح لم يُجرَ. ما يمكن قوله مقيس: ما نملكه من مخزون،
       وحجم الجمهور، ومن يجلس على الأصل الآن."""
    wash = next(t for t in P["types"] if t["key"] == "carwash")
    shop = next(t for t in P["types"] if t["key"] == "shop")
    return [
        dict(product="إكسسوارات سيارات",
             holder="مستأجرونا أنفسهم — ثم أسواق قطع الغيار خارج المحطة",
             evidence=f"{wash['leased']} مغسلة و{shop['leased']} محلاً مؤجَّرة تبيع الخدمة "
                      f"على أرضنا اليوم، ونأخذ إيجاراً لا هامشاً",
             asset=f"{wash['vacant']} مغسلة و{shop['vacant']} محلاً شاغرة",
             need="سلة منتجات وأسعار المنافس في ثلاثة أحياء · وقرار: نبيع أم نؤجّر"),
        dict(product="مساحات إعلانية",
             holder="مشغّلو الإعلان الخارجي — والمعلن يشتري منهم لا منّا",
             evidence=f"{visits_month/1e6:.2f} مليون معاملة شهرياً جمهورٌ لا يُباع، "
                      f"ووجه علبة المناديل أول سطح مسعَّر عندنا",
             asset="واجهات المظلات والمضخات وشاشات نقاط البيع — بلا جرد",
             need="جرد الأسطح لكل محطة · وسعر السوق للوجه الواحد شهرياً"),
    ]


# ═══════════════════════════════════════════════ استهداف كل منفذ
def targeting():
    """من يستهدفه كل منفذ — مقيس من مزيج الوقود والسلة والذروة وحصة المساء

       والمسح التنافسي الميداني يغطّي خمسة منافذ فقط؛ لا يُملأ الباقي بتقدير."""
    P = json.load(open(PLAN, encoding="utf-8"))
    F = {f["code"]: f for f in P.get("five", [])}
    SG = {g["seg"]: g for g in P["segments"]}
    ST = sorted(P["stations"], key=lambda x: -x["sar_total"])

    def who(x):
        w = ["شاحنات وأساطيل" if x["diesel"] >= 0.30 else
             "سكّان وعابرون" if x["diesel"] >= 0.10 else "أفراد المدينة"]
        if x["fleet"] >= 0.02: w.append("حسابات أسطول رقمية")
        if x["night"] >= 0.55: w.append("طلب مسائي غالب")
        return " · ".join(w)

    profiles = []
    for seg in SG:
        g = [x for x in ST if x["seg"] == seg]
        if not g:
            continue
        profiles.append(dict(
            cls=seg, n=len(g), mlpa=sum(x["mlpa"] for x in g),
            diesel=stx.median([x["diesel"] for x in g]),
            lpv=stx.median([x["lpv"] for x in g]),
            vpd=stx.median([x["vpd"] for x in g]),
            peak=int(stx.median([x["peak"] for x in g])),
            night=stx.median([x["night"] for x in g]),
            who=SGWHO.get(seg, "")))
    profiles.sort(key=lambda g: -g["mlpa"])

    top = [dict(code=x["code"], name=x["name"], cls=x["seg"], who=who(x),
                peak=x["peak"], sar=round(x["sar_total"]),
                scanned=x["code"] in F,
                comp_n=F[x["code"]]["n"] if x["code"] in F else 0)
           for x in ST[:5]]
    mism = [x for x in ST if (x["diesel"] >= 0.30 and x["seg"] != "خط سفر")
            or (x["diesel"] < 0.10 and x["seg"] == "خط سفر")]
    return dict(n=len(ST), profiles=profiles, top=top,
                scanned=len(F), mismatch=len(mism),
                mismatches=[dict(code=x["code"], name=x["name"], cls=x["seg"],
                                 diesel=x["diesel"], lpv=x["lpv"]) for x in mism])


def build():
    D = json.load(open(DASH, encoding="utf-8"))
    visits = sum(v["vis"] for v in D["network"].values())
    P = property_()
    out = dict(retail=retail(), corporate=corporate(), property=P,
               underwrite=underwrite(), targeting=targeting(),
               newcomers=newcomers(P, visits / 7),
               months=7, visits=visits, visits_month=visits / 7)
    json.dump(out, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    return out


if __name__ == "__main__":
    d = build()
    R, C, P, U = d["retail"], d["corporate"], d["property"], d["underwrite"]
    print(f"① الأفراد: {R['n_sample']} منافساً في العيّنة من {R['n_total']} داخل ٥ كم · "
          f"تقييمهم {R['rate_mean']:.2f}★ ونحن {R['ours_mean']:.2f}★ (فارق {R['gap_mean']:+.2f})")
    for b in R["brands"]:
        print(f"   {b['brand']:<14}{b['n']:>3} ({b['share']:>5.1%})  أقرب {b['near']:>5}م  "
              f"تقييم {b['rating']:.2f}★" if b["rating"] else "")
    print(f"   النمو من {R['growth_lo']:+.0f}٪ إلى {R['growth_hi']:+.0f}٪ والميزة ثابتة — "
          f"ارتباط النمو بفارق التقييم {R['r_gap']:+.2f} وبعدد المنافسين {R['r_n']:+.2f}")
    print(f"   شكاوانا {R['n_complaints']} · حصة الخدمة منها {R['svc_share']:.0%}")
    print(f"\n② الشركات: منصّات الأساطيل {C['plat_rev']:,.0f} ريال ({C['plat_share']:.2%} من المسجَّل) · "
          f"فاتورتها {C['plat_inv']:.0f} ريال مقابل نقدي {C['cash_inv']:.0f}")
    print(f"   الجهاز الخارجي {C['pos_rev']:,.0f} بفاتورة {C['pos_inv']:.0f} — نقاط بيع لا أسطول")
    print(f"   محفظتنا {C['our_rev']:,.0f} ريال في {C['our_vis']} معاملات")
    print(f"   قائمة الصيد: {len(C['gap'])} محطة ديزلها {C['gap_diesel']:,.0f} ريال "
          f"({C['gap_vol']:,.0f} لتر) بلا أي حضور منصّة")
    print(f"\n③ العقار: {P['n']:,} وحدة في {P['stations']} محطة · إشغال {P['occ']:.1%}")
    for t in P["types"]:
        print(f"   {t['ar']:<8}{t['n']:>6,}  إشغال {t['occ']:>6.1%}  شاغر {t['vacant']:>5,}")
    for c in P["cats"]:
        print(f"   {c['cat']:<14}{c['n']:>6,}  إشغال {c['occ']:>6.1%}  شاغر {c['vacant']:>5,}")
    print(f"\n④ اكتتاب الوحدات: {U['n']} محطة قابلة للحساب · وسيط {U['median']:,.0f} زيارة/يوم لكل وحدة شاغرة")
    print(f"   المدى {U['spread']:.0f}× — من {U['top'][0]['per']:,.0f} إلى {U['bottom'][-1]['per']:,.0f}")
    for x in U["top"][:4]:
        print(f"   {x['code']:<7}{x['name'][:20]:<22}شاغرة {x['vacant']:>3} · {x['per']:>6,.0f} زيارة/يوم لكل وحدة")
    print("   ...")
    for x in U["bottom"]:
        print(f"   {x['code']:<7}{x['name'][:20]:<22}شاغرة {x['vacant']:>3} · {x['per']:>6,.0f} زيارة/يوم لكل وحدة")
    print(f"   التغطية: {U['measured']:,} من {U['total_vacant']:,} وحدة شاغرة نملك مُدخَلها الأول ({U['measured_share']:.0%})")
    for c in U["coverage"]:
        print(f"     {c['cat']:<14}{c['measured']:>4} من {c['vacant']:>5} — {c['why']}")
