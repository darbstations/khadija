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

   ويُستخلَص ثلاثةُ مُدخَلات:
     · معدّلُ النظير   — وسيطُ لتر/يوم لمحطاتِ المدينة (ن≥٣) وإلا المنطقة
     · منحنى التدرّج   — من المحطات التي افتُتحت في أثناء ٢٠٢٦ ولها ٥ أشهرٍ فعليّة
     · مزيجُ المنتجات  — حصصُ ٩١/٩٥/الديزل الفعليّةُ لمنطقة المحطة

   يُشغَّل:  python3.13 -I data/read_new_stations.py <مسار الملف>
             (و‏-I لأنّ الملف مُرفَقٌ غيرُ موثوق)
"""
import collections
import datetime as dt
import json
import statistics as st
import sys

import openpyxl

OUT = "outlets/data/new-stations-2027.json"
DAYS = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]   # ٢٠٢٧ — فبراير ٢٨
NACT = 9          # ينا–سبت: آخرُ شهرٍ فعليٍّ في ملف ٢٠٢٦
PRODS = ("91", "95", "diesel")

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
    net = st.median(list(rate.values()))

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
        r, src, n = peer(s)
        m = mix.get(s["reg"], mnet)
        bu = {p: vol[c][p][2027] if p in vol[c] else [0.0] * 12 for p in PRODS}
        a26 = {p: vol[c][p][2026] if p in vol[c] else [0.0] * 12 for p in PRODS}
        out.append(dict(
            code=c, name=s["name"], model=s["model"], reg=s["reg"], city=s["city"],
            open=s["open"],
            # «شهرُه»: موضعُ الافتتاح نسبةً إلى يناير ٢٠٢٧، مرقَّمٌ من صفر
            off=(d.year - 2027) * 12 + d.month - 1,
            rate=round(r), src=src, n=n, mix=m, bu=bu, a26=a26,
            has26=any(sum(a26[p]) for p in PRODS)))

    D = dict(stations=out, ramp=curve, days=DAYS,
             peer_city={k: [round(st.median(v)), len(v)]
                        for k, v in city.items() if len(v) >= 3},
             peer_reg={k: [round(st.median(v)), len(v)] for k, v in reg.items()},
             peer_net=round(net), n_mature=len(rate), n_ramp=nramp,
             mix_reg=mix, mix_net=mnet)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(D, f, ensure_ascii=False, indent=1)

    bu = sum(sum(sum(s["bu"][p]) for p in PRODS) for s in out)
    print(f"✓ {OUT}")
    print(f"  محطات {len(out)} · ناضجةٌ للمعايرة {len(rate)} · "
          f"افتتاحاتٌ قيس بها التدرّج {nramp}")
    print(f"  منحنى التدرّج: {' · '.join(f'{x:.2f}' for x in curve)}")
    print(f"  وسيطُ الشبكة {net:,.0f} ل/يوم · موازنةُ الـ٣٨ {bu / 1e6:.1f}م لتر")
    return D


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit("الاستخدام: python3.13 -I data/read_new_stations.py <ملف الموازنة>")
    build(sys.argv[1])
