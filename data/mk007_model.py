# -*- coding: utf-8 -*-
"""محطة العمرة الجديدة MK007 — التشخيص قبل المستهدف

   المصدر: تقرير كاش إن اليومي للمحطة · ٢٦٤ يوماً (١ يناير – ٢١ سبتمبر ٢٠٢٦)
   ١٬١٤٠٬٢٠٨ عملية · محفوظ في outlets/data/mk007-daily.csv بصف لكل يوم.
   يوم ٢١ سبتمبر ناقص فيُستبعد من كل متوسط، ويُذكر أنه استُبعد.

   والترتيب في هذا الملف هو ترتيب القاعدة نفسها:
   ① ما الذي تغيّر ومتى بالضبط   → periods() و breakpoint()
   ② ما تصنيف السبب               → classify()
   ③ بأي رقم نحكم ومتى            → target() و cost()

   يُشغَّل من جذر المستودع:  PYTHONPATH=data python3 data/mk007_model.py
"""
import csv, json, statistics as stx

import sales_plan as S

DAILY = "outlets/data/mk007-daily.csv"
META = "outlets/data/mk007-meta.json"
HOURS = "outlets/data/network-hours.csv"
PLAN = "data/sales-plan.json"
COMPLAINTS = "outlets/data/complaints.csv"
STAFF = "outlets/data/station-staff.csv"
TANKS = "outlets/data/tanks.csv"
OUT = "data/mk007.json"

CODE = "MK007"
SEG = "حيوية"                 # تصنيف الشركة من analytics.csv
PROD = ["g91", "g95", "diesel"]
PRODN = {"g91": "بنزين ٩١", "g95": "بنزين ٩٥", "diesel": "ديزل"}
PAY = ["cash", "visa", "network", "syarah", "petro", "smartcard", "other"]
PAYN = {"cash": "نقدي", "visa": "بطاقة", "network": "شبكة",
        "syarah": "سيارة", "petro": "بترومين", "smartcard": "بطاقة ذكية",
        "other": "أخرى"}
QN = {1: "الربع الأول", 2: "الربع الثاني", 3: "الربع الثالث — حتى ٢٠ سبتمبر"}
MONTHN = ["", "يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
          "يوليو", "أغسطس", "سبتمبر"]


def _f(v, d=0.0):
    try: return float(v)
    except (TypeError, ValueError): return d


# ═══════════════════════════════════════════════════ القراءة
def read():
    """الأيام الكاملة وحدها — اليوم الناقص يُحمل جانباً ولا يدخل متوسطاً"""
    rows, partial = [], None
    for r in csv.DictReader(open(DAILY, encoding="utf-8")):
        d = dict(date=r["date"], tx=int(_f(r["tx"])), sales=_f(r["sales"]),
                 litres=_f(r["litres"]), ge60=int(_f(r["ge60"])),
                 eq60=int(_f(r["eq60"])), b2b=int(_f(r["b2b"])),
                 b2b_sales=_f(r["b2b_sales"]), dur=_f(r["dur"]),
                 b=[int(_f(r["b%d" % i])) for i in range(6)],
                 h=[int(_f(r["h%02d" % i])) for i in range(24)],
                 hs=[_f(r["hs%02d" % i]) for i in range(24)],
                 f={p: [int(_f(r[p + "_tx"])), _f(r[p + "_sales"]),
                        _f(r[p + "_litres"])] for p in PROD},
                 p={p: [int(_f(r[p + "_tx"])), _f(r[p + "_sales"])] for p in PAY})
        d["month"] = int(d["date"][5:7])
        d["q"] = (d["month"] - 1) // 3 + 1
        if int(_f(r["partial"])): partial = d
        else: rows.append(d)
    return rows, partial


def agg(rows):
    """مجمّع واحد يُستعمل للربع وللشهر ولنافذة الحملة — فلا تتعدّد التعاريف"""
    n = len(rows) or 1
    t = dict(days=len(rows), tx=0, sales=0.0, litres=0.0, ge60=0, eq60=0,
             b2b=0, b2b_sales=0.0)
    f = {p: [0, 0.0, 0.0] for p in PROD}
    p = {k: [0, 0.0] for k in PAY}
    b = [0] * 6; h = [0] * 24; hs = [0.0] * 24; durw = 0.0
    for r in rows:
        for k in ("tx", "sales", "litres", "ge60", "eq60", "b2b", "b2b_sales"):
            t[k] += r[k]
        for k in PROD:
            for i in range(3): f[k][i] += r["f"][k][i]
        for k in PAY:
            for i in range(2): p[k][i] += r["p"][k][i]
        for i in range(6): b[i] += r["b"][i]
        for i in range(24): h[i] += r["h"][i]; hs[i] += r["hs"][i]
        durw += r["dur"] * r["tx"]
    tx = t["tx"] or 1
    t.update(vpd=t["tx"] / n, lpd=t["litres"] / n, spd=t["sales"] / n,
             lpv=t["litres"] / tx, inv=t["sales"] / tx,
             ge60_pct=t["ge60"] / tx, b2b_pct=t["b2b"] / tx,
             dur=durw / tx, price=t["sales"] / (t["litres"] or 1),
             diesel=f["diesel"][2] / (t["litres"] or 1),
             prod={k: dict(name=PRODN[k], tx=v[0], sales=v[1], litres=v[2],
                           share=v[2] / (t["litres"] or 1),
                           lpv=v[2] / (v[0] or 1),
                           price=v[1] / (v[2] or 1)) for k, v in f.items()},
             pay={k: dict(name=PAYN[k], tx=v[0], sales=v[1],
                          share=v[1] / (t["sales"] or 1)) for k, v in p.items()},
             bands=[x / tx for x in b], bands_n=b,
             h=[x / n for x in h], hs=[x / n for x in hs],
             morning=sum(h[0:12]) / n, evening=sum(h[12:24]) / n)
    t["eve_share"] = t["evening"] / (t["morning"] + t["evening"] or 1)
    return t


# ═══════════════════════════════════════════════════ ① ما الذي تغيّر ومتى
def periods(rows):
    qs = {q: agg([r for r in rows if r["q"] == q]) for q in (1, 2, 3)}
    ms = {m: agg([r for r in rows if r["month"] == m])
          for m in sorted({r["month"] for r in rows})}
    return qs, ms


def breakpoint(rows):
    """الانحدار ليلةً واحدة منافسةٌ أو عطل · والانزلاق ستة أسابيع شيءٌ آخر.
       نقيس سرعة التغيّر: أكبر فرق بين متوسط ١٤ يوماً وما قبله."""
    ser = [(r["date"], r["litres"], r["tx"],
            r["f"]["diesel"][2] / (r["litres"] or 1),
            r["p"]["cash"][1] / (r["sales"] or 1)) for r in rows]
    out = []
    for i in range(14, len(ser) - 13):
        a = stx.mean(x[1] for x in ser[i - 14:i])
        b = stx.mean(x[1] for x in ser[i:i + 14])
        out.append((ser[i][0], b / a - 1))
    worst = min(out, key=lambda x: x[1])
    # مدى الانزلاق: من القمة إلى القاع بمتوسط ١٤ يوماً
    roll = [(ser[i][0], stx.mean(x[1] for x in ser[i:i + 14]))
            for i in range(len(ser) - 13)]
    hi = max(roll, key=lambda x: x[1]); lo = min(roll, key=lambda x: x[1])
    span = (list(dict(roll)).index(lo[0]) - list(dict(roll)).index(hi[0]))
    return dict(worst_date=worst[0], worst_pct=worst[1],
                peak_date=hi[0], peak_lpd=hi[1],
                trough_date=lo[0], trough_lpd=lo[1],
                drop=lo[1] / hi[1] - 1, span_days=span,
                recovered=roll[-1][1], recovery=roll[-1][1] / lo[1] - 1,
                vs_peak=roll[-1][1] / hi[1] - 1)


def baselines(rows, ms):
    """أي خط أساس تختار يقرّر أترى أزمةً أم استواءً — فتُعرض الثلاثة معاً"""
    dates = sorted(r["date"] for r in rows)
    by = {r["date"]: r for r in rows}
    def win(a, b):
        g = [by[d] for d in dates if a <= d <= b]
        return stx.mean(r["litres"] for r in g), stx.mean(r["tx"] for r in g)
    jan_l, jan_t = win("2026-01-01", "2026-01-31")
    pk_l, pk_t = win("2026-03-03", "2026-03-16")
    now_l, now_t = win("2026-09-07", "2026-09-20")
    # نظائر مكة — يناير مقابل يوليو من ملف الشبكة (وهو كل ما نملكه عنها)
    peers = []
    for r in csv.DictReader(open("outlets/data/network-sales.csv", encoding="utf-8")):
        if r["region"].strip() != "مكة": continue
        j, u = _f(r["jan_lpd"]), _f(r["jul_lpd"])
        if j <= 0 or u <= 0: continue
        peers.append(dict(code=r["code"], name=r["name"], jan=j, jul=u,
                          chg=u / j - 1,
                          diesel=_f(r["diesel_vol"]) / (_f(r["volume"]) or 1)))
    peers.sort(key=lambda x: -x["jan"])
    return dict(
        rows=[dict(name="يناير — قبل الموسم", lpd=jan_l, vpd=jan_t),
              dict(name="ذروة مارس ٣–١٦", lpd=pk_l, vpd=pk_t),
              dict(name="آخر أربعة عشر يوماً", lpd=now_l, vpd=now_t)],
        vs_jan_l=now_l / jan_l - 1, vs_jan_t=now_t / jan_t - 1,
        vs_peak_l=now_l / pk_l - 1, vs_peak_t=now_t / pk_t - 1,
        peers=peers, peers_n=len(peers),
        peers_med=stx.median([p["chg"] for p in peers]) if peers else 0.0)


def handover(H, q3):
    """أكبر فجوة ساعةٍ واحدة تقع خارج «ساعات الذروة» — لأن التعريف يستعمل
       شكلنا المنخفض نفسه. فتُعرض على حدة ولا تُدسّ في الرقم المؤكَّد."""
    r23 = H["rows"][23]
    worst = sorted([r for r in H["rows"] if not r["peak"]],
                   key=lambda r: -r["gap"])[:3]
    return dict(row=r23, worst=worst,
                extra=sum(x["gap"] for x in worst),
                note="شكاوى المحطة تذكر تسليم الورديات صراحةً سبباً للتأخير عند منتصف الليل")


def channels(rows):
    """متى ظهرت كل وسيلة دفع ومتى اختفت — تغيّر النظام لا يُقرأ حملةً"""
    out = []
    for k in PAY:
        seen = [r["date"] for r in rows if r["p"][k][0] > 0]
        if not seen: continue
        n = sum(r["p"][k][0] for r in rows)
        out.append(dict(key=k, name=PAYN[k], first=seen[0], last=seen[-1],
                        days=len(seen), tx=n,
                        sales=sum(r["p"][k][1] for r in rows)))
    return sorted(out, key=lambda x: -x["tx"])


# ═══════════════════════════════════════════════════ الساعة والوردية
def ref_profile():
    """شكل ساعات تصنيف «حيوية» — وسيط ٢٥ محطة من ملف ساعات الشبكة"""
    plan = json.load(open(PLAN, encoding="utf-8"))
    seg = {s["code"]: s["seg"] for s in plan["stations"]}
    prof = []
    for r in csv.DictReader(open(HOURS, encoding="utf-8")):
        if seg.get(r["code"]) != SEG: continue
        v = [_f(r["h%02d" % h]) for h in range(24)]
        t = sum(v)
        if t: prof.append([x / t for x in v])
    return [stx.median([p[h] for p in prof]) for h in range(24)], len(prof)


def hourly(q3):
    """الفجوة تُحتسب حيث نحن دون شكل تصنيفنا — لا حيث الطلب نفسه أقلّ.
       وساعات الذروة وحدها فجوةٌ قابلة للدفاع، وما عداها يحتاج إثباتاً."""
    ref, nref = ref_profile()
    vpd = q3["vpd"]
    share = [x / vpd for x in q3["h"]]
    top = set(sorted(range(24), key=lambda h: -share[h])[:8])
    rows = []
    for h in range(24):
        gap = max(0.0, (ref[h] - share[h])) * vpd
        rows.append(dict(hour=h, actual=q3["h"][h], sales=q3["hs"][h],
                         share=share[h], ref=ref[h], expect=ref[h] * vpd,
                         gap=gap, peak=h in top,
                         shift="صباحية" if h < 12 else "مسائية"))
    peak = sum(r["gap"] for r in rows if r["peak"])
    quiet = sum(r["gap"] for r in rows if not r["peak"])
    return dict(ref_n=nref, rows=rows, peak_gap=peak, quiet_gap=quiet,
                gap_m=sum(r["gap"] for r in rows if r["hour"] < 12),
                gap_e=sum(r["gap"] for r in rows if r["hour"] >= 12),
                peak_m=sum(r["gap"] for r in rows if r["peak"] and r["hour"] < 12),
                peak_e=sum(r["gap"] for r in rows if r["peak"] and r["hour"] >= 12),
                busiest=max(rows, key=lambda r: r["actual"]),
                widest=max(rows, key=lambda r: r["gap"]))


def shifts(q3, H):
    """حِمل الوردية مقيسٌ لا مُقدَّر: معاملات الوردية ÷ عمّالها.
       وعدد العمّال من ثلاثة ملفات لا يتفق — فيُعرض الثلاثة ولا يُنتقى واحد."""
    st = {r["code"]: r for r in csv.DictReader(open(STAFF, encoding="utf-8"))}
    s = st.get(CODE, {})
    dw, ew = int(_f(s.get("day_workers"), 0)), int(_f(s.get("eve_workers"), 0))
    tot = dw + ew
    m, e = q3["morning"], q3["evening"]
    bal = (m + e) / (tot or 1)
    need_m, need_e = m / bal, e / bal
    # سعة الخدمة النظرية: العامل الواحد ٣٬٦٠٠ ثانية ÷ زمن الخدمة المقيس
    cap = 3600.0 / (q3["dur"] or 1)
    peak_hour = max(range(12, 24), key=lambda h: q3["h"][h])
    return dict(day_w=dw, eve_w=ew, total_w=tot,
                morning=m, evening=e, eve_share=q3["eve_share"],
                eve_staff_share=ew / (tot or 1),
                load_m=m / (dw or 1), load_e=e / (ew or 1),
                ratio=(e / (ew or 1)) / (m / (dw or 1)),
                balanced=bal, need_m=need_m, need_e=need_e,
                move=round(dw - need_m / 1.0) if need_m < dw else 0,
                move_exact=dw - need_m,
                cap_worker=cap, dur=q3["dur"],
                peak_hour=peak_hour, peak_txn=q3["h"][peak_hour],
                peak_cap=cap * ew,
                utilisation=q3["h"][peak_hour] / (cap * ew or 1))


# ═══════════════════════════════════════════════════ ② تصنيف السبب
def classify(qs, H, sh, cmp_, tank):
    """مصفوفة القاعدة نفسها — ولكل سطر علامته في بيانات هذه المحطة وحدها"""
    q1, q3 = qs[1], qs[3]
    out = []

    def row(kind, sign, verdict, act, owner, gate=False):
        out.append(dict(kind=kind, sign=sign, verdict=verdict, act=act,
                        owner=owner, gate=gate))

    row("وصول", "هبوط مفاجئ مع ثبات اللتر/الزيارة",
        "لا — الهبوط انزلاقٌ في ستة أسابيع لا ليلةً، واللتر/الزيارة هبط معه",
        "لا إجراء", "—")
    row("بيانات", "وسيلة الدفع غير مسجَّلة · فجوة مطابقة",
        f"نعم — «غير معروف» {q3['pay']['other']['share']*100:.2f}٪ فقط، لكن ثلاث "
        f"قنوات جديدة دخلت في أبريل ونسبة النقدي هبطت "
        f"{(q1['pay']['cash']['share']-q3['pay']['cash']['share'])*100:.0f} نقطة ولم تعُد",
        "توثيق تغيير القنوات في تعريف خط الأساس قبل أي مقارنة ربعية",
        "تقنية المعلومات + المالية", gate=True)
    row("عقد", "العمولة أو النسبة تلتهم الهامش",
        "خارج نطاق هذا الملف — المحطة استثمارية في تقرير المالية وإيجارية في ملف "
        "العمليات وتشغيلية في سجلّ الوحدات",
        "توحيد نموذج العمل قبل احتساب أي عائد",
        "المالية + العمليات", gate=True)
    row("إمداد", "تغطية خزان دون أربعة أيام",
        f"نعم — تغطية الديزل {tank['diesel_days']:.1f} يوم وسعة المحطة كلها "
        f"{tank['all_days']:.1f} يوم عند معدّل الربع الثالث",
        "جدول توريد مُثبَّت قبل أي مستهدف يرفع الطلب",
        "سلسلة الإمداد", gate=True)
    row("رقابة", "نقدي المحطة شاذّ عن وسيط الشبكة",
        f"يحتاج فحصاً — النقدي {q3['pay']['cash']['share']*100:.0f}٪ بعد أن كان "
        f"{q1['pay']['cash']['share']*100:.0f}٪، والشكاوى تذكر سحب مبالغ زائدة",
        "مطابقة نقاط البيع بالصندوق لشهر واحد", "المراجعة الداخلية")
    row("عمالة", "عبء وردية يفوق أختها",
        f"نعم وبوضوح — العامل المسائي يخدم {sh['load_e']:.0f} معاملة مقابل "
        f"{sh['load_m']:.0f} للصباحي ({sh['ratio']:.2f}×)، وكل فجوة الذروة "
        f"({H['peak_e']:.0f} معاملة/يوم) في المسائية وحدها",
        f"نقل {abs(sh['move']):.0f} عمّال من الصباحية إلى المسائية — بلا توظيف",
        "العمليات")
    row("مبيعات", "ما بقي بعد الفرز",
        "ما بقي: فجوة ساعات الذروة المسائية بعد إغلاق البوابات أعلاه",
        "مستهدف معاملات للوردية المسائية — لا حملة عامة",
        "التجاري")
    return out


def tank_cover(q3):
    t = {r["code"]: r for r in csv.DictReader(open(TANKS, encoding="utf-8-sig"))}
    r = t.get(CODE, {})
    cap = {"g91": _f(r.get("cap_91")), "g95": _f(r.get("cap_95")),
           "diesel": _f(r.get("cap_diesel"))}
    days = {k: cap[k] / (q3["prod"][k]["litres"] / q3["days"] or 1) for k in cap}
    return dict(cap=cap, days=days, diesel_days=days["diesel"],
                all_days=sum(cap.values()) / (q3["lpd"] or 1),
                lpd={k: q3["prod"][k]["litres"] / q3["days"] for k in cap})


def complaints():
    rows = list(csv.DictReader(open(COMPLAINTS, encoding="utf-8-sig")))
    mine = [r for r in rows if r.get("station") == CODE]
    cat = {}
    for r in mine: cat[r.get("category") or "—"] = cat.get(r.get("category") or "—", 0) + 1
    netcat = {}
    for r in rows: netcat[r.get("category") or "—"] = netcat.get(r.get("category") or "—", 0) + 1
    slow = cat.get("بطء الخدمة", 0)
    return dict(n=len(mine), net=len(rows), share=len(mine) / (len(rows) or 1),
                cats=sorted(cat.items(), key=lambda x: -x[1]),
                slow=slow, net_slow=netcat.get("بطء الخدمة", 0),
                slow_share=slow / (netcat.get("بطء الخدمة", 0) or 1))


# ═══════════════════════════════════════════════════ ③ المستهدف وكلفته
def target(q3, H, sh):
    """يُطالَب بالمعاملات — وحجم التعبئة يُفسَّر بالمزيج ولا يُطالَب به.
       والقِسمة على الوردية والعامل تُشتقّ من الفجوة ذاتها لا من رغبة."""
    d = q3["diesel"]
    blend = S.MARGIN_PETROL * (1 - d) + S.MARGIN_DIESEL * d
    gap = H["peak_gap"]
    close = S.CLOSE
    add = gap * close
    litres = add * q3["lpv"]
    return dict(diesel=d, blend=blend, lpv=q3["lpv"], inv=q3["inv"],
                gap_peak=gap, gap_quiet=H["quiet_gap"], close=close,
                add_txn=add, add_txn_full=gap,
                add_litres=litres, add_litres_full=gap * q3["lpv"],
                add_sales=add * q3["inv"],
                cm_day=litres * blend, cm_year=litres * blend * 365,
                cm_year_full=gap * q3["lpv"] * blend * 365,
                per_txn_year=q3["lpv"] * blend * 365,
                base_vpd=q3["vpd"], base_lpd=q3["lpd"],
                tgt_vpd=q3["vpd"] + add, tgt_lpd=q3["lpd"] + litres,
                eve_share=H["peak_e"] / (H["peak_gap"] or 1),
                add_eve=add * H["peak_e"] / (H["peak_gap"] or 1),
                per_worker=add * H["peak_e"] / (H["peak_gap"] or 1) / (sh["eve_w"] or 1),
                hours=[dict(hour=r["hour"], gap=r["gap"], add=r["gap"] * close)
                       for r in H["rows"] if r["peak"] and r["gap"] > 0])


def cost(q3, T):
    """ثلاثة أُسُس للكلفة ولكلٍّ قرارٌ يخصّه — ولا يُخلط أساسٌ بقرار غيره"""
    L = T["add_litres"]
    bases = [
        ("حدّية — لتر إضافي على طاقم قائم",
         0.0, "لا أجر ولا مرفق إضافياً: النقص طاقةُ خدمة لا طاقة ضخّ",
         "يصلح لقرار: هل نغلق الفجوة بالطاقم الحالي؟"),
        ("تشغيل المحطة — محمَّل بالكامل",
         S.OPEX_NET, "٩٫١٥ هللة من قائمة الدخل — أجور ومرافق وصيانة",
         "يصلح لقرار: هل نفتح وردية أو نوظّف؟"),
        ("المحمَّل الكامل — مع مصاريف المركز",
         S.OPEX_NET + S.LOADED, "يُضاف ٤٫٤١ هللة بيعاً وتسويقاً وإدارة وفوائد",
         "يصلح لقرار: هل تستحق المحطة رأس مال جديداً؟"),
    ]
    out = []
    for name, cpl, note, use in bases:
        out.append(dict(base=name, cpl=cpl, note=note, use=use,
                        cost_day=L * cpl, cost_year=L * cpl * 365,
                        net_day=T["cm_day"] - L * cpl,
                        net_year=(T["cm_day"] - L * cpl) * 365,
                        cpl_net=T["blend"] - cpl))
    # آلية الهدية: الكلفة تقع على كل فاتورة مؤهَّلة لا على الإضافية وحدها
    #   والمقيس عندنا شريحة «٦٠ ريالاً فأكثر»، وعتبة الخطة ٥٠ — فيُعرض حدّان
    elig = q3["ge60_pct"]
    elig40 = sum(q3["bands"][2:])          # ٤٠ ريالاً فأكثر — الحدّ الأعلى
    all_day = q3["vpd"] * elig * S.BOX
    all40 = q3["vpd"] * elig40 * S.BOX
    inc_day = T["add_txn"] * elig * S.BOX
    cm_station = q3["lpd"] * T["blend"] * 365
    return dict(bases=out, box=S.BOX, threshold=S.THRESHOLD,
                elig_pct=elig, elig_day=q3["vpd"] * elig,
                elig40_pct=elig40, gift40_year=all40 * 365,
                gift_all_day=all_day, gift_all_year=all_day * 365,
                gift_inc_day=inc_day, gift_inc_year=inc_day * 365,
                cm_year=T["cm_year"], ratio=all_day * 365 / (T["cm_year"] or 1),
                cm_station=cm_station, of_station=all_day * 365 / cm_station,
                of_station40=all40 * 365 / cm_station,
                breakeven_litres=all_day / (T["blend"] or 1),
                breakeven_pct=all_day / (T["blend"] or 1) / (q3["lpd"] or 1),
                breakeven_txn=all_day / (T["lpv"] * T["blend"] or 1),
                unmeasured=[
                    ("رسوم البطاقات", "غير مقيسة",
                     f"{(1-q3['pay']['cash']['share'])*100:.0f}٪ من المبيعات بطاقةٌ اليوم "
                     f"بعد أن كانت {100-58.5:.0f}٪ — والرسم غير معروف لنا",
                     "المالية — جدول رسوم مدى/فيزا"),
                    ("الفروقات والفاقد", "غير مقيسة",
                     "لا جرد خزانات في بياناتنا — والفرق يظهر هامشاً لا كلفة",
                     "العمليات — جرد شهري"),
                    ("النقل والتوريد", "غير مقيسة",
                     "من يتحمّل أجرة الناقلة غير محسوم في أي ملف لدينا",
                     "سلسلة الإمداد"),
                ])


def campaigns(rows):
    """ما لم يُصمَّم قياسه قبلاً يُنسب أثره بعداً إلى أي سبب.
       فنقيسه هنا بثلاث نوافذ — قبل وأثناء وبعد — ونقارن بتشتّت خط الأساس."""
    meta = json.load(open(META, encoding="utf-8"))
    by = {r["date"]: r for r in rows}
    dates = sorted(by)
    idx = {d: i for i, d in enumerate(dates)}
    # تشتّت طبيعي: انحراف نوافذ ١٤ يوماً عن وسيطها
    w = [stx.mean(by[d]["tx"] for d in dates[i:i + 14])
         for i in range(0, len(dates) - 14, 7)]
    noise = stx.pstdev(w) / stx.median(w)
    out = []
    for c in meta["campaigns"]:
        i0, i1 = idx.get(c["from"]), idx.get(c["to"])
        if i0 is None or i1 is None: continue
        pre = [by[d] for d in dates[max(0, i0 - 14):i0]]
        run = [by[d] for d in dates[i0:i1 + 1]]
        post = [by[d] for d in dates[i1 + 1:i1 + 15]]
        if not (pre and run and post): continue
        def mor(g): return stx.mean(sum(r["h"][5:12]) for r in g)
        def txn(g): return stx.mean(r["tx"] for r in g)
        def g60(g): return sum(r["ge60"] for r in g) / sum(r["tx"] for r in g)
        metric = "الفترة الصباحية ٠٥–١١" if c["focus"] == "morning" else "حصة فواتير ٦٠ ريالاً فأكثر"
        if c["focus"] == "morning":
            a, b, cc = mor(pre), mor(run), mor(post)
            unit = "معاملة/يوم"
        else:
            a, b, cc = g60(pre) * 100, g60(run) * 100, g60(post) * 100
            unit = "٪"
        lift = b / a - 1 if a else 0
        after = cc / a - 1 if a else 0
        if c["focus"] == "morning":
            verdict = ("لا يُنسب للحملة — ما بعدها أعلى منها" if cc > b
                       else "أثرٌ ضمن التشتّت الطبيعي" if abs(lift) <= noise
                       else "أثرٌ يفوق التشتّت")
        else:
            verdict = ("أثرٌ أثناء التنفيذ بلا بقاء بعده" if cc <= a < b
                       else "أثرٌ ضمن التشتّت الطبيعي" if abs(lift) <= noise
                       else "أثرٌ يفوق التشتّت")
        out.append(dict(name=c["name"], offer=c["offer"], mech=c["mech"],
                        frm=c["from"], to=c["to"], days=len(run),
                        metric=metric, unit=unit,
                        pre=a, run=b, post=cc, lift=lift, after=after,
                        txn_pre=txn(pre), txn_run=txn(run), txn_post=txn(post),
                        txn_lift=txn(run) / txn(pre) - 1,
                        verdict=verdict))
    return dict(noise=noise, rows=out)


def unknowns():
    return [
        ("جدول الورديات المعتمد", "العمليات",
         "٢٠ عاملاً في ملف الطاقم و٢٦ في تقرير العمّال و٣٢ في ملف العمليات",
         "ملزم قبل أي مستهدف وردية"),
        ("نموذج عمل المحطة", "المالية",
         "استثماري · إيجاري · تشغيلي — ثلاثة ملفات وثلاث إجابات",
         "يغيّر أساس الكلفة كلّه"),
        ("رسوم البطاقات", "المالية",
         "٥٥٪ من المبيعات بطاقة ولا نعرف رسمها",
         "يُقتطع من هامش كل لتر إضافي"),
        ("جرد الخزانات", "العمليات",
         "لا مطابقة بين المُورَّد والمُباع في بياناتنا",
         "يفصل فجوة المطابقة عن فجوة البيع"),
        ("جدول هامش المنتجات", "المالية",
         "١٢٫٦٦ و٤٫٤٧ هللة استُنتجا بانحدار ١٩ محطة لا بجدول معتمد",
         "الديزل ٣٣٪ هنا فأثر الخطأ مباشر"),
        ("سعر لتر الديزل لعملاء الجملة", "التجاري",
         f"عملاء الشركات {'—'} من المعاملات، وشروطهم غير مسجَّلة عندنا",
         "يفسّر تذبذب حصة الديزل"),
    ]


# ═══════════════════════════════════════════════════ البناء
def build():
    rows, partial = read()
    qs, ms = periods(rows)
    q3 = qs[3]
    H = hourly(q3)
    sh = shifts(q3, H)
    tank = tank_cover(q3)
    cmp_ = campaigns(rows)
    T = target(q3, H, sh)
    C = cost(q3, T)
    meta = json.load(open(META, encoding="utf-8"))
    D = dict(
        code=CODE, name=meta["name"], seg=SEG, source=meta["source"],
        rows=meta["rows"], dupes=meta["dupes"], skipped=meta["skipped"],
        first=rows[0]["date"], last=rows[-1]["date"], days=len(rows),
        partial=meta["partial"], bands=meta["bands"],
        quarters={str(k): v for k, v in qs.items()},
        qnames={str(k): QN[k] for k in qs},
        months={str(k): v for k, v in ms.items()},
        monthnames=MONTHN,
        year=agg(rows),
        brk=breakpoint(rows), channels=channels(rows),
        base=baselines(rows, ms), handover=handover(H, q3),
        hourly=H, shifts=sh, tank=tank, complaints=complaints(),
        causes=classify(qs, H, sh, cmp_, tank),
        target=T, cost=C, campaigns=cmp_, unknowns=unknowns(),
        const=dict(margin_petrol=S.MARGIN_PETROL, margin_diesel=S.MARGIN_DIESEL,
                   opex=S.OPEX_NET, loaded=S.LOADED, box=S.BOX,
                   threshold=S.THRESHOLD, close=S.CLOSE),
    )
    json.dump(D, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    return D


if __name__ == "__main__":
    D = build()
    q = D["quarters"]; T = D["target"]; H = D["hourly"]; sh = D["shifts"]
    print(f"{D['code']} · {D['name']} · {D['seg']} — {D['days']} يوماً كاملاً "
          f"({D['first']} → {D['last']}) · {D['rows']:,} عملية")
    for k in ("1", "2", "3"):
        r = q[k]
        print(f"  {D['qnames'][k]:<26} {r['vpd']:>7,.0f} معاملة/يوم · "
              f"{r['lpd']:>9,.0f} لتر/يوم · {r['lpv']:>5.2f} لتر/معاملة · "
              f"ديزل {r['diesel']*100:>4.1f}٪ · نقدي {r['pay']['cash']['share']*100:>4.1f}٪")
    b = D["brk"]; bl = D["base"]
    print(f"\n  القمة {b['peak_date']} ({b['peak_lpd']:,.0f} لتر/يوم) ← "
          f"القاع {b['trough_date']} ({b['trough_lpd']:,.0f}) = {b['drop']*100:+.1f}٪ "
          f"في {b['span_days']} يوماً · والآن {b['recovered']:,.0f} "
          f"({b['vs_peak']*100:+.1f}٪ عن القمة)")
    print(f"  وخط الأساس يقرّر الحكم: مقابل يناير {bl['vs_jan_l']*100:+.1f}٪ لتراً "
          f"و{bl['vs_jan_t']*100:+.1f}٪ معاملةً · مقابل ذروة مارس {bl['vs_peak_l']*100:+.1f}٪ "
          f"· ونظائر مكة ({bl['peers_n']}) وسيطها {bl['peers_med']*100:+.1f}٪")
    ho = D["handover"]
    print(f"  وأكبر فجوة ساعةٍ ({ho['row']['hour']:02d}:٠٠) "
          f"{ho['row']['gap']:.0f} معاملة/يوم — خارج الرقم المؤكَّد بحكم التعريف")
    print(f"  فجوة الذروة {H['peak_gap']:.0f} معاملة/يوم — منها "
          f"{H['peak_e']:.0f} في المسائية و{H['peak_m']:.0f} في الصباحية")
    print(f"  العامل المسائي {sh['load_e']:.0f} معاملة مقابل {sh['load_m']:.0f} "
          f"للصباحي ({sh['ratio']:.2f}×) · استغلال ساعة الذروة "
          f"{sh['utilisation']*100:.0f}٪")
    print(f"  المستهدف {T['add_txn']:.0f} معاملة/يوم = {T['cm_year']:,.0f} ريال/سنة "
          f"هامش مساهمة")
    c = D["cost"]
    print(f"  كلفة الهدية على كل فاتورة مؤهَّلة {c['gift_all_year']:,.0f} ريال/سنة "
          f"= {c['ratio']:.1f}× العائد")
    print("كُتب", OUT)
