# -*- coding: utf-8 -*-
"""تكلفة اللتر الواحد — لا قيمته

   القيمة تُقاس بالهامش، والتكلفة تُقاس بما نُخرجه من جيبنا لنبيع لتراً:
   شراؤه، ثم تشغيل المحطة التي باعته، ثم ما يُحمَّل عليه من مصاريف المركز.

   مصدران، ولكلٍّ نطاقه — ولا يُخلطان:
   ① قائمة الدخل · يوليو ٢٠٢٦ تراكمي · ١٤١ محطة · ٦٦١٫٤ مليون لتر
      تعطي التكلفة الكاملة بأربع طبقات، وبتفصيل نموذج العمل.
   ② تقرير ربحية الربع الثاني · ٢١ محطة مشغّلة · ٤٩٫١ مليون لتر
      وحده يفصّل داخل «مصاريف التشغيل»: أجور · مرافق · ما بقي.

   يُشغَّل من جذر المستودع:  PYTHONPATH=data python3 data/litre_cost.py
"""
import csv, json

COSTS = "outlets/data/q2-costs.csv"      # أجور ومرافق كل محطة — ر١ مقابل ر٢
Q2 = "outlets/data/q2-stations.csv"
OUT = "data/litre-cost.json"

# ── ① قائمة الدخل · يناير–يوليو ٢٠٢٦ · بآلاف الريالات (صفحة ٨ من تقرير الأداء)
#    الأعمدة: استثماري · إيجاري · تشغيلي · امتياز
MODELS = ["استثماري", "إيجاري", "تشغيلي", "امتياز"]
PL = {
    "stations":  [6, 17, 26, 92],
    "litres_m":  [91.8, 77.7, 141.2, 350.8],       # مليون لتر
    "revenue":   [174309, 146564, 268915, 583790],
    "cogs":      [163544, 137777, 252931, 579654],  # تكلفة المبيعات — شراء الوقود
    "fuel_cm":   [10765, 8787, 15984, 4136],        # هامش مساهمة الوقود
    "rent_cm":   [6427, -7141, 1248, 60],           # هامش مساهمة الإيجار
    "opex":      [10167, 4378, 13896, 144],         # مصاريف تشغيل المحطات
    "selling":   [597, 597, 717, 478],              # بيع وتسويق — محمّل
    "admin":     [2921, 2921, 3505, 2337],          # عمومية وإدارية — محمّل
    "interest":  [761, 761, 913, 609],              # فوائد القروض — محمّل
    "net":       [2722, -7038, -1869, 609],         # صافي الربح
}

# ── معدّل الأجر في التقرير: ثابت لكل عامل، يُضرب في عدد العمّال ولا يُقاس
#    تحقّق على ٢١ محطة: الأجور ÷ عدد العمّال = ٣٬٥٠٠٫٠ ريالاً بلا استثناء واحد
WAGE_RATE = 3500.0


def _f(v, d=0.0):
    try: return float(v)
    except (TypeError, ValueError): return d


def income_statement():
    """① التكلفة الكاملة بأربع طبقات — الشبكة كلها وكل نموذج عمل"""
    n = len(MODELS)
    tot = {k: (sum(v) if k != "litres_m" else sum(v)) for k, v in PL.items()}
    L = tot["litres_m"] * 1e6                       # لترات الشبكة

    def h(thousands, litres_m):                     # ألف ريال ← هللة لكل لتر
        return thousands * 1e3 / (litres_m * 1e6) * 100

    layers = [
        ("شراء الوقود", "cogs", "ما ندفعه لأرامكو — لا نملك فيه قراراً"),
        ("تشغيل المحطة", "opex", "أجور ومرافق وصيانة — هنا وحدها تُضبط التكلفة"),
        ("بيع وتسويق", "selling", "محمّل من المركز"),
        ("عمومية وإدارية", "admin", "محمّل من المركز"),
        ("فوائد القروض", "interest", "محمّل من المركز"),
    ]
    stack = [dict(layer=t, key=k, note=note,
                  total=tot[k], cpl=h(tot[k], tot["litres_m"]),
                  by_model=[h(PL[k][i], PL["litres_m"][i]) for i in range(n)])
             for t, k, note in layers]
    cost_cpl = sum(s["cpl"] for s in stack)
    rev_cpl = h(tot["revenue"], tot["litres_m"])
    rent_cpl = h(tot["rent_cm"], tot["litres_m"])
    return dict(
        models=MODELS, stations=PL["stations"], litres_m=PL["litres_m"],
        total_litres=L, total_stations=sum(PL["stations"]),
        stack=stack, cost_cpl=cost_cpl, rev_cpl=rev_cpl, rent_cpl=rent_cpl,
        net_cpl=h(tot["net"], tot["litres_m"]),
        fuel_cm_cpl=h(tot["fuel_cm"], tot["litres_m"]),
        by_model=[dict(model=MODELS[i], stations=PL["stations"][i],
                       litres_m=PL["litres_m"][i],
                       rev=h(PL["revenue"][i], PL["litres_m"][i]),
                       cogs=h(PL["cogs"][i], PL["litres_m"][i]),
                       cm=h(PL["fuel_cm"][i], PL["litres_m"][i]),
                       opex=h(PL["opex"][i], PL["litres_m"][i]),
                       loaded=h(PL["selling"][i] + PL["admin"][i] + PL["interest"][i],
                                PL["litres_m"][i]),
                       net=h(PL["net"][i], PL["litres_m"][i]))
                  for i in range(n)],
        totals=tot)


def operating_detail():
    """② داخل «تشغيل المحطة»: أجور · مرافق · ما بقي — من تقرير الربع الثاني"""
    C = {r["code"]: r for r in csv.DictReader(open(COSTS, encoding="utf-8"))}
    Q = {r["code"]: r for r in csv.DictReader(open(Q2, encoding="utf-8"))}
    rows, tw, tu, tvol, tsales, tstaff = [], 0.0, 0.0, 0.0, 0.0, 0.0
    for c, q in Q.items():
        vol = _f(q["q2_vol"]); w = _f((C.get(c) or {}).get("q2_wages"))
        u = _f((C.get(c) or {}).get("q2_util")); n = _f(q["q2_staff"])
        if not vol or not w:
            continue
        tw += w; tu += u; tvol += vol; tsales += _f(q["q2_sales"]); tstaff += n
        rows.append(dict(code=c, name=q["name"], vol=vol, staff=n,
                         wages=w, util=u,
                         wage_cpl=w / vol * 100, util_cpl=u / vol * 100,
                         litres_per_worker=vol / n if n else 0))
    rows.sort(key=lambda r: -r["wage_cpl"])
    # المصاريف التشغيلية الكلية من التقرير نفسه — لاشتقاق «ما بقي»
    OPEX_TOTAL, GP_TOTAL, VOL_TOTAL, SALES_TOTAL = 2385991.0, 5442225.0, 49079686.0, 91979684.0
    other = OPEX_TOTAL - tw - tu
    return dict(
        n=len(rows), rows=rows, staff=tstaff,
        wages=tw, util=tu, other=other, opex=OPEX_TOTAL,
        vol=VOL_TOTAL, sales=SALES_TOTAL,
        cogs=SALES_TOTAL - GP_TOTAL,
        cpl=dict(cogs=(SALES_TOTAL - GP_TOTAL) / VOL_TOTAL * 100,
                 wages=tw / VOL_TOTAL * 100,
                 util=tu / VOL_TOTAL * 100,
                 other=other / VOL_TOTAL * 100,
                 opex=OPEX_TOTAL / VOL_TOTAL * 100,
                 revenue=SALES_TOTAL / VOL_TOTAL * 100,
                 gross=GP_TOTAL / VOL_TOTAL * 100,
                 net=(GP_TOTAL - OPEX_TOTAL) / VOL_TOTAL * 100),
        wage_share=tw / OPEX_TOTAL, util_share=tu / OPEX_TOTAL,
        wage_rate=WAGE_RATE,
        # تحذير لا يُطمس: الأجر في التقرير معدَّل ثابت لا كشف رواتب
        caveat=("الأجور في التقرير = عدد العمّال × ٣٬٥٠٠ ريالاً للربع بلا استثناء واحد "
                "في ٢١ محطة — أي ١٬١٦٧ ريالاً شهرياً للعامل. هذا معدَّل معياري لا كشف "
                "رواتب فعلي، فتكلفة العمالة الحقيقية أعلى، وكل رقم تحتها يتحرّك معها."))


def build():
    out = dict(pl=income_statement(), op=operating_detail())
    json.dump(out, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    return out


if __name__ == "__main__":
    d = build(); P, O = d["pl"], d["op"]
    print(f"① الشبكة كلها — {P['total_stations']} محطة · "
          f"{P['total_litres']/1e6:,.1f} مليون لتر (يناير–يوليو ٢٠٢٦)\n")
    print(f"   {'الطبقة':<18}{'ألف ريال':>12}{'هللة/لتر':>11}   {'ملاحظة'}")
    for s in P["stack"]:
        print(f"   {s['layer']:<18}{s['total']:>12,.0f}{s['cpl']:>11.2f}   {s['note']}")
    print(f"   {'─'*58}")
    print(f"   {'إجمالي التكلفة':<18}{'':>12}{P['cost_cpl']:>11.2f}")
    print(f"   {'سعر البيع المحقَّق':<18}{'':>12}{P['rev_cpl']:>11.2f}")
    print(f"   {'مساهمة الإيجار':<18}{'':>12}{P['rent_cpl']:>11.2f}")
    print(f"   {'صافي اللتر':<18}{'':>12}{P['net_cpl']:>11.2f}   "
          f"{'خسارة' if P['net_cpl'] < 0 else 'ربح'}\n")
    print(f"   {'النموذج':<12}{'محطات':>7}{'م لتر':>8}{'البيع':>9}{'الشراء':>9}"
          f"{'الهامش':>8}{'التشغيل':>9}{'المحمّل':>9}{'الصافي':>9}")
    for m in P["by_model"]:
        print(f"   {m['model']:<12}{m['stations']:>7}{m['litres_m']:>8.1f}{m['rev']:>9.2f}"
              f"{m['cogs']:>9.2f}{m['cm']:>8.2f}{m['opex']:>9.2f}{m['loaded']:>9.2f}"
              f"{m['net']:>9.2f}")
    c = O["cpl"]
    print(f"\n② محطات التشغيل — {O['n']} محطة · {O['vol']/1e6:,.1f} مليون لتر (الربع الثاني)\n")
    for k, ar in (("cogs", "شراء الوقود"), ("wages", "أجور العمّال"),
                  ("util", "كهرباء ومياه وهاتف"), ("other", "بقية التشغيل")):
        print(f"   {ar:<20}{c[k]:>8.2f} هللة/لتر")
    print(f"   {'─'*30}")
    print(f"   {'إجمالي التكلفة':<20}{c['cogs']+c['opex']:>8.2f}")
    print(f"   {'سعر البيع':<20}{c['revenue']:>8.2f}")
    print(f"   {'صافي اللتر':<20}{c['net']:>8.2f}")
    print(f"\n   الأجور {O['wage_share']:.0%} من المصاريف التشغيلية · والمرافق {O['util_share']:.0%}")
    print(f"   ⚠ {O['caveat']}")
    print(f"\n   أعلى خمس محطات كلفةَ عمالة لكل لتر:")
    for r in O["rows"][:5]:
        print(f"     {r['code']:<7}{r['name'][:16]:<18}{r['wage_cpl']:>6.2f} هللة · "
              f"{r['staff']:>3.0f} عاملاً · {r['litres_per_worker']:>9,.0f} لتر للعامل")
    print(f"   وأدناها:")
    for r in O["rows"][-3:]:
        print(f"     {r['code']:<7}{r['name'][:16]:<18}{r['wage_cpl']:>6.2f} هللة · "
              f"{r['staff']:>3.0f} عاملاً · {r['litres_per_worker']:>9,.0f} لتر للعامل")
