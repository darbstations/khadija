#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
تحقق وظيفي من صحة دوال اللوحة باستخدام محرّك Excel حقيقي (pycel).

يبني نسخة اختبار من المصنّف، يحقن بيانات تجريبية في صفحتَي الانقطاعات والردود،
يحسب القيم المتوقّعة بشكل مستقل، ثم يقيّم خلايا اللوحة فعلياً ويؤكّد التطابق.

التشغيل:
    pip install openpyxl pycel
    python3 scripts/verify_formulas.py        # يخرج برمز 0 عند النجاح، 1 عند الفشل
"""
import calendar
import datetime
import sys
import tempfile
import warnings

import openpyxl

warnings.filterwarnings("ignore")

WB = "workbook/khadija-supply-chain-kpis-2026.xlsx"
DASH = "📊 ملخص آلي + KPIs"
PRODUCT_POINTS = 300  # قيمة خلية الإدخال F38

# (بدء, انتهاء) — مدة الانقطاع بالأيام = الفرق ؛ يناير 2+1=3 ، فبراير 3+2=5 ، مارس لا شيء
OUTAGES = [
    (datetime.date(2026, 1, 5), datetime.date(2026, 1, 7), "محطة 1", "بنزين 91"),
    (datetime.date(2026, 1, 10), datetime.date(2026, 1, 11), "محطة 2", "ديزل"),
    (datetime.date(2026, 2, 3), datetime.date(2026, 2, 6), "محطة 1", "بنزين 95"),
    (datetime.date(2026, 2, 20), datetime.date(2026, 2, 22), "محطة 3", "ديزل"),
]
# (تاريخ, محطة, منتج, صهريج, كمية, فاقد) — يناير 80,130 ، فبراير 50,70,90 ، مارس لا شيء
RETURNS = [
    (datetime.date(2026, 1, 8), "محطة 1", "بنزين 91", "ص1", 33000, 80),
    (datetime.date(2026, 1, 20), "محطة 2", "ديزل", "ص2", 30000, 130),
    (datetime.date(2026, 2, 5), "محطة 1", "بنزين 95", "ص3", 33000, 50),
    (datetime.date(2026, 2, 15), "محطة 2", "ديزل", "ص4", 32000, 70),
    (datetime.date(2026, 2, 25), "محطة 3", "بنزين 91", "ص5", 33000, 90),
]


def build_test_file(path: str) -> None:
    wb = openpyxl.load_workbook(WB)
    o = wb["🛢️ انقطاعات الوقود"]
    for i, (b, e, st, p) in enumerate(OUTAGES):
        o.cell(5 + i, 2).value, o.cell(5 + i, 3).value = b, e
        o.cell(5 + i, 4).value, o.cell(5 + i, 5).value = st, p
    r = wb["💧 الردود والفاقد"]
    for i, (d, st, p, tk, q, l) in enumerate(RETURNS):
        r.cell(5 + i, 2).value, r.cell(5 + i, 3).value = d, st
        r.cell(5 + i, 4).value, r.cell(5 + i, 5).value = p, tk
        r.cell(5 + i, 6).value, r.cell(5 + i, 7).value = q, l
    wb.save(path)


def expected_kpi1(m: int) -> float:
    s = sum((e - b).days for b, e, *_ in OUTAGES if b.year == 2026 and b.month == m)
    return 1 - s / (PRODUCT_POINTS * calendar.monthrange(2026, m)[1])


def expected_kpi6(m: int) -> float:
    ls = [l for d, *_, l in RETURNS if d.month == m]
    return sum(ls) / len(ls) if ls else 0


def main() -> int:
    from pycel import ExcelCompiler

    path = tempfile.mktemp(suffix=".xlsx")
    build_test_file(path)
    exc = ExcelCompiler(path)

    cases = []  # (cell, label, expected)
    for col, m, name in [("G", 1, "يناير"), ("H", 2, "فبراير"), ("I", 3, "مارس")]:
        cases.append((f"{col}7", f"توفر الوقود {name}", round(expected_kpi1(m), 6)))
        cases.append((f"{col}18", f"فواقد الوقود {name}", expected_kpi6(m)))
    # خلايا الحالة + مؤشر غير مُعدّل للتأكد أن اللوحة لم تنكسر
    status = [("F7", "حالة توفر الوقود", "✅ +0.5%"),
              ("F18", "حالة الفواقد", "✅ -12.5%"),
              ("G9", "KPI2 عدد انقطاعات يناير", 2),
              ("I9", "KPI2 عدد انقطاعات مارس", 0)]

    ok = True
    print(f"{'cell':<6}{'label':<26}{'actual':>12}{'expected':>12}  ")
    for cell, label, exp in cases + status:
        val = exc.evaluate(f"'{DASH}'!{cell}")
        if isinstance(exp, (int, float)) and not isinstance(exp, bool):
            try:
                passed = abs(float(val) - exp) < 1e-4
            except (TypeError, ValueError):
                passed = False
        else:
            passed = str(val) == str(exp)
        ok &= passed
        shown = round(val, 6) if isinstance(val, float) else val
        print(f"{cell:<6}{label:<26}{str(shown):>12}{str(exp):>12}  {'✅' if passed else '❌'}")

    print("\nRESULT:", "✅ ALL FORMULAS CORRECT & FUNCTIONAL" if ok else "❌ FAILURES DETECTED")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
