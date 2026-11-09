# -*- coding: utf-8 -*-
"""قراءة ملف «sales & Rent Budget 2027.xlsx» إلى سلاسلَ موثّقةِ المصدر

   الملف ليس في المستودع (يُرسَل من المالية)، فهذا الملفّ يحفظ *من أين*
   يُقرأ كل رقم، ومخرجُه `outlets/data/budget-2027.json` هو ما يُبنى عليه.

   وفيه مصيدتان لا يُقرأ الملف بدونهما:

   ① **ورقة الوقود فيها صفّان لكل محطة.** ٣٩٨ صفَّ محطة مفاتيحُها `Code-U`
      وصيغتُها `=SUMIFS(N:N,$J:$J,$K<row>)`، وتحتها ١٬١٦١ صفَّ تفصيلٍ لكل
      محطة‑منتج. فمن جمع الاثنين ضاعف اللترات. والمرجع صفُّ الإجمالي في
      الورقة نفسها: `N417` = `N450` = ٨٩٬٥٠٢٬٢٩٥ لتراً في يناير.

   ② **ورقة الإيجار فيها ١٨٣ صفَّ محطة** (`Status-Station = "Station"`)
      إيرادُها صفر، فلا تُضاعِف شيئاً — لكنها تُفسد عدّ الوحدات.

   وحدُّ الفعلي ليس حيث يبدو:
     · الوقود  — يناير–سبتمبر فعليّ، وأكتوبر–ديسمبر من محرّك الملف
                 (`Q4-2026 growth vs Q3-2026 = 0.03`).
     · الإيجار — يناير–يوليو قيمٌ ملصوقة، وأغسطس فيه ٣٩٧ خليةً صيغتُها
                 «= يوليو»، وسبتمبر–ديسمبر منقولةٌ من الشهر السابق. وإجمالي
                 الملف الفعلي `AZ` يجمع يناير–أغسطس، والتحصيل موسومٌ
                 `Act` حتى أغسطس و`FC` بعده.

   يُشغَّل:  python3 data/read_budget_2027.py <مسار الملف>
"""
import collections
import json
import re
import sys

import openpyxl

OUT = "outlets/data/budget-2027.json"

# ── أعمدة ورقة الوقود «GC│Liters & Value subledger» (صفر‑الأساس)
GC = dict(model=2, region=3, city=4, name=7, link=9, code=10,
          l26=13, l27=25, v26=38, v27=50, scount26=63, newold=125)
# ── أعمدة ورقة الإيجار «GD│Rent roll (1) subledger» (صفر‑الأساس)
GD = dict(model=1, station=2, code=5, unit=8, stastatus=14, accept=15,
          utype=16, area=18, stdann=20,
          gross=39,      # AN–AY · Actual-Gross Revenue
          credit=53,     # CD–CO · Credit Notes (المجموع)
          net=95,        # CR–DC · Net Revenue
          coll=109,      # DF–DQ · Collection (Act حتى أغسطس)
          b27=135,       # EF–EQ · Budget-Gross Revenue 2027
          b27basis=132,  # EC    · Budget-27 basis
          p26a=228,      # HU–IF · P-26-Revenue      ← مخطط «أ»
          p26b=408)      # OS–PD · P-26-Rental Income ← مخطط «ب»

FUEL_ACT = 9      # آخر شهرٍ فعليّ في الوقود (سبتمبر)
RENT_ACT = 8      # آخر شهرٍ يدخل إجمالي الملف الفعلي في الإيجار (أغسطس)


def num(v):
    return float(v) if isinstance(v, (int, float)) else 0.0


def is_formula(v):
    return isinstance(v, str) and v.startswith("=")


def read(path):
    wv = openpyxl.load_workbook(path, data_only=True, read_only=True)
    wf = openpyxl.load_workbook(path, data_only=False, read_only=True)
    out = {"source": path.split("/")[-1],
           "fuel_actual_months": FUEL_ACT, "rent_actual_months": RENT_ACT}

    # ═══════════ الوقود — صفوف المحطات وحدها
    V = list(wv["GC│Liters & Value subledger"]
             .iter_rows(min_row=1, max_col=130, values_only=True))
    F = list(wf["GC│Liters & Value subledger"]
             .iter_rows(min_row=1, max_col=130, values_only=True))
    S = [V[i - 1] for i in range(8, len(V) + 1)
         if is_formula(F[i - 1][GC["l26"]])
         and F[i - 1][GC["l26"]].startswith("=SUMIFS(N:N")]
    out["fuel_station_rows"] = len(S)
    for key, off in (("l26", GC["l26"]), ("l27", GC["l27"]),
                     ("v26", GC["v26"]), ("v27", GC["v27"])):
        out["fuel_" + key] = [sum(num(r[off + k]) for r in S) for k in range(12)]
        out["fuel_" + key + "_n"] = [sum(1 for r in S if num(r[off + k]) > 0)
                                     for k in range(12)]
    # المرجع: صفُّ الإجمالي في الورقة نفسها
    tot = [i for i in range(8, len(V) + 1)
           if is_formula(F[i - 1][GC["l26"]])
           and F[i - 1][GC["l26"]].startswith("=SUM(N8")]
    out["fuel_check_jan"] = num(V[tot[-1] - 1][GC["l26"]]) if tot else None

    def l26(r): return [num(r[GC["l26"] + k]) for k in range(12)]
    def l27(r): return [num(r[GC["l27"] + k]) for k in range(12)]
    out["fuel_sold_actual"] = sum(1 for r in S if sum(l26(r)[:FUEL_ACT]) > 0)
    out["fuel_any_26"] = sum(1 for r in S if sum(l26(r)) > 0)
    out["fuel_has_27"] = sum(1 for r in S if sum(l27(r)) > 0)
    out["fuel_new_27"] = sum(1 for r in S
                             if sum(l26(r)) == 0 and sum(l27(r)) > 0)
    out["fuel_all_actual_months"] = sum(
        1 for r in S if all(l26(r)[k] > 0 for k in range(FUEL_ACT)))
    # النموذج والمنطقة — للمحطات التي باعت فعلاً
    for fld, col in (("model", GC["model"]), ("region", GC["region"])):
        g = collections.defaultdict(lambda: [0, 0.0])
        for r in S:
            if sum(l26(r)[:FUEL_ACT]) <= 0:
                continue
            k = str(r[col] or "").strip()
            g[k][0] += 1
            g[k][1] += sum(l26(r)[:FUEL_ACT])
        out["fuel_by_" + fld] = {k: v for k, v in g.items()}
    # معدّل المحطة الجديدة مقابل القائمة في خطة ٢٧
    days = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

    def rate27(r):
        d = sum(days[k] for k in range(12) if l27(r)[k] > 0)
        return sum(l27(r)) / d if d else 0.0
    new = sorted(rate27(r) for r in S
                 if sum(l26(r)) == 0 and sum(l27(r)) > 0)
    old = sorted(rate27(r) for r in S
                 if sum(l26(r)) > 0 and sum(l27(r)) > 0)
    med = lambda a: a[len(a) // 2] if a else 0.0
    out["fuel_rate_new_27"] = dict(n=len(new), median=med(new), min=min(new or [0]))
    out["fuel_rate_old_27"] = dict(n=len(old), median=med(old))

    # ═══════════ الإيجار — صفوف الوحدات وحدها
    R = [r for r in wv["GD│Rent roll (1) subledger"]
         .iter_rows(min_row=9, max_col=460, values_only=True)
         if r[GD["code"]] and str(r[GD["code"]]).strip()
         and not str(r[GD["station"]] or "").strip().lower().startswith("total")
         and str(r[GD["stastatus"]] or "").strip() != "Station"]
    out["rent_unit_rows"] = len(R)
    for key in ("gross", "credit", "net", "coll", "b27", "p26a", "p26b"):
        off = GD[key]
        out["rent_" + key] = [sum(num(r[off + k]) for r in R) for k in range(12)]
        out["rent_" + key + "_n"] = [sum(1 for r in R if num(r[off + k]) > 0)
                                     for k in range(12)]
    acc = lambda r: str(r[GD["accept"]] or "").strip().lower()
    sta = lambda r: str(r[GD["stastatus"]] or "").strip()
    bill = [r for r in R if num(r[GD["gross"] + RENT_ACT - 1]) > 0]
    out["rent_billing"] = len(bill)
    out["rent_flag"] = dict(collections.Counter(acc(r) for r in R))
    out["rent_flag_occupied_billing"] = sum(
        1 for r in R if acc(r) == "occupied"
        and num(r[GD["gross"] + RENT_ACT - 1]) > 0)
    out["rent_flag_occupied_idle"] = sum(
        1 for r in R if acc(r) == "occupied"
        and num(r[GD["gross"] + RENT_ACT - 1]) == 0)
    out["rent_billing_unflagged"] = sum(1 for r in bill if acc(r) != "occupied")
    out["rent_billing_area"] = sum(num(r[GD["area"]]) for r in bill)
    out["rent_billing_stations"] = len({str(r[GD["code"]]).strip() for r in bill})
    out["rent_billing_by_type"] = dict(collections.Counter(
        str(r[GD["utype"]] or "").strip() for r in bill))
    out["rent_billing_by_model"] = dict(collections.Counter(
        str(r[GD["model"]] or "").strip() for r in bill))
    out["rent_units_by_model"] = dict(collections.Counter(
        str(r[GD["model"]] or "").strip() for r in R))
    # مصدر موازنة ٢٧: وحدةٌ قائمة أم شاغرةٌ في محطةٍ لم تُفتح
    g = collections.defaultdict(lambda: [0, 0.0])
    for r in R:
        b = sum(num(r[GD["b27"] + k]) for k in range(12))
        if b <= 0:
            continue
        k = str(r[GD["b27basis"]] or "").strip() or "(فارغ)"
        g[k][0] += 1
        g[k][1] += b
    out["rent_b27_basis"] = {k: v for k, v in g.items()}
    out["rent_b27_vacant_notactive"] = sum(
        1 for r in R if acc(r) == "vacant" and sta(r) == "Not Active"
        and sum(num(r[GD["b27"] + k]) for k in range(12)) > 0)
    out["rent_pool_active_vacant"] = sum(
        1 for r in R if acc(r) == "vacant" and sta(r) == "Active")
    out["rent_pool_in_b27"] = sum(
        1 for r in R if acc(r) == "vacant" and sta(r) == "Active"
        and sum(num(r[GD["b27"] + k]) for k in range(12)) > 0)
    # وحداتٌ تفوتر اليوم وموازنتُها صفر
    zero = [r for r in bill if sum(num(r[GD["b27"] + k]) for k in range(12)) == 0]
    out["rent_billing_zero_b27"] = dict(
        units=len(zero),
        annual=sum(num(r[GD["gross"] + RENT_ACT - 1]) for r in zero) * 12,
        stations=dict(collections.Counter(
            str(r[GD["station"]] or "").strip()[:30] for r in zero).most_common(6)))
    # السعر المعياري حيث يوجد مستأجر
    both = [r for r in bill if num(r[GD["stdann"]]) > 0]
    out["rent_vs_standard"] = dict(
        units=len(both), of=len(bill),
        actual=sum(num(r[GD["gross"] + RENT_ACT - 1]) for r in both) * 12,
        standard=sum(num(r[GD["stdann"]]) for r in both),
        no_standard=len(bill) - len(both),
        no_standard_annual=sum(num(r[GD["gross"] + RENT_ACT - 1])
                               for r in bill if num(r[GD["stdann"]]) == 0) * 12)

    # ═══════════ السجل — جدول الافتتاح
    st = collections.Counter()
    sched = collections.Counter()
    for r in wv["1│Register"].iter_rows(min_row=15, max_col=26, values_only=True):
        if not r[0] or not str(r[0]).strip():
            continue
        s = str(r[7] or "").strip()
        st[s] += 1
        m = re.search(r"opens\s+([A-Za-z]{3})-(\d\d)", s)
        if m:
            sched[f"{m.group(1)}-{m.group(2)}"] += 1
    out["register_total"] = sum(st.values())
    out["register_existing"] = st.get("Existing", 0)
    out["register_schedule"] = dict(sched)
    out["register_opens_26"] = sum(v for k, v in sched.items() if k.endswith("-26"))
    out["register_opens_27"] = sum(v for k, v in sched.items() if k.endswith("-27"))
    return out


def main():
    if len(sys.argv) < 2:
        sys.exit("الاستعمال: python3 data/read_budget_2027.py <ملف الموازنة.xlsx>")
    d = read(sys.argv[1])
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=1)
    print(f"✓ {OUT}")
    print(f"  الوقود: {d['fuel_station_rows']} صفَّ محطة · "
          f"يناير {d['fuel_l26'][0]:,.0f} لتراً "
          f"(مرجع الورقة {d['fuel_check_jan']:,.0f})")
    print(f"  الإيجار: {d['rent_unit_rows']} صفَّ وحدة · "
          f"تفوتر {d['rent_billing']}")
    print(f"  السجل: {d['register_total']} محطة · "
          f"قائمة {d['register_existing']} · "
          f"تفتح ٢٦ {d['register_opens_26']} · ٢٧ {d['register_opens_27']}")


if __name__ == "__main__":
    main()
