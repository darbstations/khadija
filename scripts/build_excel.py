#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
مولّد النموذج المالي لبرنامج تانكي (نسخة v5) — بمعادلات حقيقية مترابطة.

ينطلق من الملف المصدر في data/tanki_model_source.xlsx ويطبّق:
  1) معدلات الكسب الموصى بها (50/75/100)
  2) هوامش درب الفعلية (13 هللة/لتر)
  3) إصلاح أخطاء معادلات عمود المؤشرات في قسم المحطات
  4) تحديث جداول المقارنة الثابتة في الملخص
  5) كتلة مدخلات "الاقتصاد الموسّع" (الديزل + الشركاء)
  6) ورقة جديدة «الاقتصاد الموسّع» بمعادلات LIVE تشير لخلايا السيناريوهات
  7) تفعيل إعادة الحساب الكامل عند الفتح

كل القيم المحسوبة في الورقة الجديدة معادلات حقيقية (=...) لا أرقاماً ثابتة.
"""
import os
import warnings
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

warnings.filterwarnings("ignore")

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SRC = os.path.join(ROOT, "data", "tanki_model_source.xlsx")
OUT = os.path.join(ROOT, "tanki_financial_model_v5.xlsx")

SC = "السيناريوهات"  # اسم ورقة السيناريوهات (للمراجع)

YELLOW = PatternFill("solid", fgColor="FFF6D26B")
HEADER = PatternFill("solid", fgColor="FF16223C")
RTL = Alignment(horizontal="right", vertical="center", readingOrder=2, wrap_text=True)
THIN = Side(style="thin", color="FF23314F")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def main() -> None:
    wb = openpyxl.load_workbook(SRC, data_only=False)
    S = wb[SC]

    # 1) معدلات الكسب الموصى بها — تنتشر لكل الأوراق عبر المعادلات
    S["C7"], S["C8"], S["C9"] = 50, 75, 100
    # 2) هوامش درب الفعلية (~13 هللة/لتر)
    S["C26"], S["C27"] = 13, 13

    # 3) إصلاح + إعادة محاذاة عمود المؤشرات (كل خلية مدمجة F#:G# → نكتب على F#)
    S["F23"] = '=IF(C23+C24+C25>=150,"✅ منطقي","⚠️ تحقق")'
    S["F24"] = ""
    S["F25"] = ""
    S["F26"] = '=IF(C26>=5,"✅ معقول","⚠️ منخفض")'
    S["F27"] = '=IF(C27>=5,"✅ معقول","⚠️ منخفض")'
    S["F28"] = '=IF(C28>=1,"✅ معقول","⚠️ ضيق")'
    S["F29"] = ""
    S["F30"] = '=IF(C29+C30=1,"✅ صحيح","🔴 خطأ المجموع")'   # الإصلاح (كان C30+C31)
    S["F31"] = '=IF(C31>=50000,"✅ ممتاز",IF(C31>=30000,"⚠️ متوسط","🔴 منخفض"))'
    S["F32"] = ""
    S["A33"] = ("💡 الفلسفة · نقاط مُضخّمة 116x–232x من ساسكو لإحساس بالثراء "
                "+ قيمة استبدال واضحة (10,000 نقطة = ريال)")

    # 4) الملخص التنفيذي — الخلايا النصية الثابتة
    M = wb["الملخص التنفيذي"]
    M["A6"], M["D6"], M["G6"] = "0.5% من كل ريال إنفاق", "0.75% للأعضاء النشطين", "1% للأعضاء المميزين"
    M["D16"], M["E16"], M["F16"], M["G16"] = "50 نقطة", "75 نقطة", "100 نقطة", "116x - 232x"
    M["D18"], M["E18"], M["F18"], M["G18"] = "0.50%", "0.75%", "1.00%", "1.2x - 2.3x"
    M["D19"], M["E19"], M["F19"] = "5,000 نقطة", "7,500 نقطة", "10,000 نقطة"

    # 5) كتلة مدخلات الاقتصاد الموسّع (الديزل + الشركاء) — خلايا صفراء قابلة للتعديل
    S["A35"] = "🆕 الاقتصاد الموسّع — مدخلات"
    S["A35"].font = Font(bold=True, color="FFFFD54A")
    ext_inputs = [
        (36, "سعر لتر الديزل (ريال)", 1.79, "يناير 2026"),
        (37, "هامش الديزل (هللة/لتر)", 4, "ضعيف جداً"),
        (38, "كسب الديزل (نقطة/ريال)", 25, "كاش باك 0.25%"),
        (39, "كسب الشركاء (نقطة/ريال)", 400, "~4% ممولة من الشريك"),
        (40, "مساهمة الشريك لدرب (هللة/ريال)", 1, "ما يدفعه الشريك مقابل عملاء درب"),
        (41, "هدية الترحيب الموصى بها (نقطة)", 50000, "= 5 ريال تكلفة فعلية"),
    ]
    for r, label, val, desc in ext_inputs:
        S.cell(r, 1, "+")
        S.cell(r, 2, label).alignment = RTL
        c = S.cell(r, 3, val)
        c.fill = YELLOW
        c.font = Font(bold=True)
        S.cell(r, 4, desc).alignment = RTL

    # 6) ورقة جديدة «الاقتصاد الموسّع» — معادلات حقيقية تشير لخلايا السيناريوهات
    if "الاقتصاد الموسّع" in wb.sheetnames:
        del wb["الاقتصاد الموسّع"]
    E = wb.create_sheet("الاقتصاد الموسّع")
    E.sheet_view.rightToLeft = True

    def put(coord, value, bold=False, fill=None, money=False):
        cell = E[coord]
        cell.value = value
        cell.alignment = RTL
        cell.font = Font(bold=bold, color="FFE8EEFC")
        if fill:
            cell.fill = fill
        if money:
            cell.number_format = "#,##0.00"
        return cell

    put("A1", "درب · محرك الاقتصاد الموسّع — الديزل والشركاء (معادلات حيّة)", bold=True)
    put("A2", "كل القيم أدناه معادلات تشير لخلايا ورقة السيناريوهات — غيّر هناك وتتحدث هنا")

    # ----- الديزل -----
    put("A4", "⛽ الديزل", bold=True, fill=HEADER)
    diesel = [
        ("A5", "البند", "B5", "القيمة", "C5", "الوحدة"),
    ]
    put("A5", "البند", bold=True); put("B5", "القيمة", bold=True); put("C5", "الوحدة", bold=True)
    rows_d = [
        (6, "هامش الديزل/لتر", f"={SC}!C37", "هللة/لتر"),
        (7, "هامش الديزل %", f"={SC}!C37/({SC}!C36*100)", "%"),
        (8, "كسب الديزل", f"={SC}!C38", "نقطة/ريال"),
        (9, "كاش باك الديزل %", f"={SC}!C38/{SC}!C6", "%"),
        (10, "تكلفة درب على الديزل", f"={SC}!C38/{SC}!C6*100", "هللة/ريال"),
        (11, "صافي هامش درب على الديزل", f"={SC}!C37/{SC}!C36-({SC}!C38/{SC}!C6*100)", "هللة/ريال"),
    ]
    for r, label, formula, unit in rows_d:
        put(f"A{r}", label)
        put(f"B{r}", formula, bold=True, money=True)
        put(f"C{r}", unit)
    E["B7"].number_format = "0.00%"
    E["B9"].number_format = "0.00%"
    put("A12", "الجدوى", bold=True)
    put("B12", '=IF(B11>0.5,"✅ مربح",IF(B11>0,"⚠️ هامش ضيق","🔴 خسارة"))', bold=True)

    # ----- الشركاء -----
    put("A14", "☕ الشركاء (كافيه/مطعم)", bold=True, fill=HEADER)
    put("A15", "البند", bold=True); put("B15", "القيمة", bold=True); put("C15", "الوحدة", bold=True)
    rows_p = [
        (16, "كسب العميل عند الشريك", f"={SC}!C39", "نقطة/ريال"),
        (17, "كاش باك العميل %", f"={SC}!C39/{SC}!C6", "%"),
        (18, "تكلفة درب (الشريك يموّل)", "=0", "هللة/ريال"),
        (19, "مساهمة الشريك لدرب", f"={SC}!C40", "هللة/ريال"),
        (20, "صافي لدرب من الشريك", f"={SC}!C40-B18", "هللة/ريال"),
    ]
    for r, label, formula, unit in rows_p:
        put(f"A{r}", label)
        put(f"B{r}", formula, bold=True, money=True)
        put(f"C{r}", unit)
    E["B17"].number_format = "0.00%"

    # ----- المقارنة: ريال وقود مقابل ريال شريك (أثره على درب) -----
    put("A22", "🔬 أثر الريال على درب: وقود مقابل شريك", bold=True, fill=HEADER)
    put("A23", "البند", bold=True); put("B23", "ريال بنزين", bold=True); put("C23", "ريال شريك", bold=True)
    put("A24", "دخل/ربح درب (هللة)")
    put("B24", f"={SC}!C26/{SC}!C11", bold=True, money=True)   # ربح بيع الوقود
    put("C24", f"={SC}!C40", bold=True, money=True)            # مساهمة الشريك
    put("A25", "تكلفة الولاء على درب (هللة)")
    put("B25", f"={SC}!C7/{SC}!C6*100", bold=True, money=True)  # تكلفة كسب الأبيض
    put("C25", "=0", bold=True, money=True)
    put("A26", "الصافي على درب (هللة/ريال)")
    put("B26", "=B24-B25", bold=True, money=True)
    put("C26", "=C24-C25", bold=True, money=True)
    put("A27", "الخلاصة", bold=True)
    put("B27", '=IF(B26>0,"✅ تكلفة احتفاظ","🔴 خسارة")', bold=True)
    put("C27", '=IF(C26>0,"✅ ربح صافٍ","—")', bold=True)

    put("A29", "💡 الوقود يجذب (هامش رفيع) والشركاء يموّلون ويربّحون — أولوية درب: توقيع الشركاء")

    # عرض الأعمدة
    for col, w in {"A": 34, "B": 16, "C": 16}.items():
        E.column_dimensions[col].width = w

    # 7) إعادة حساب كاملة عند الفتح
    try:
        wb.calculation.fullCalcOnLoad = True
    except Exception:
        pass

    wb.save(OUT)
    print("تم الحفظ:", OUT)


if __name__ == "__main__":
    main()
