# -*- coding: utf-8 -*-
"""جدول مقارنة: المخطط · الفعلي · الموازنة — من ملف موازنة ٢٠٢٧

   ورقةٌ واحدة، وأربعةُ أقسام، وكلُّ نسبةٍ فيها صيغةٌ حيّة لا رقمٌ مكتوب:

   ① الإيجار · الإيراد المفوتَر — والمخطط فيه نسختان في الملف نفسه
   ② الإيجار · عدد الوحدات المفوتَرة — وهو الرقم الذي يُحكم به لا الإيراد
   ③ الوقود · اللترات والمحطات — ولا مخطط ٢٠٢٦ في الملف أصلاً
   ④ المحطات · سجلُّ الافتتاح مقابل ما باع فعلاً

   والمصدر `outlets/data/budget-2027.json` يُنتجه `data/read_budget_2027.py`،
   وفيه حدُّ الفعلي لكل جانب: الوقود تسعةُ أشهر، والإيجار ثمانية.

   يُشغَّل من جذر المستودع:  PYTHONPATH=data python3 data/build_budget_compare.py
"""
import json

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import brand as B
import worker_model as W

SRC = "outlets/data/budget-2027.json"
OUT = "docs/مقارنة-المخطط-والفعلي-والموازنة.xlsx"

MONEY = '#,##0;[Red](#,##0)'
NUM = '#,##0'
NUM1 = '#,##0.0'
NUM3 = '#,##0.000'
PCT = '0%'
PCT1 = '0.0%'

M = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
     "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"]
NC = 9                      # B..J
HEADF = PatternFill("solid", fgColor=B.BGRAY)
CALCF = PatternFill("solid", fgColor=B.T_NEUTRAL)
FCF = PatternFill("solid", fgColor=B.T_GOLD)      # شهرٌ غير فعليّ
BANDF = PatternFill("solid", fgColor=B.T_BAND)


def txt(ws, r, c, v, *, fmt=None, font=None, fill=None, align=None):
    x = ws.cell(r, c, v)
    x.font = font or W.BLACK
    x.alignment = align or W.CTR
    x.border = W.BOX
    if fill:
        x.fill = fill
    if fmt:
        x.number_format = fmt
    return x


def band(ws, r, text):
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=NC)
    c = ws.cell(r, 2, text)
    c.font = Font(name=B.FONT, size=12, bold=True, color=B.ORANGE)
    c.fill = PatternFill("solid", fgColor=B.T_ORANGE)
    c.alignment = W.RGT
    c.border = W.BOX
    for j in range(3, NC + 1):
        cc = ws.cell(r, j)
        cc.border = W.BOX
        cc.fill = PatternFill("solid", fgColor=B.T_ORANGE)
    ws.row_dimensions[r].height = 22
    return r + 1


def note(ws, r, text, *, tone=None, h=28):
    fl = PatternFill("solid", fgColor={"gold": B.T_GOLD, "bad": B.T_BAD,
                                       "good": B.T_GOOD}.get(tone, B.BG))
    col = {"gold": B.D_GOLD, "bad": B.D_BAD, "good": B.D_GOOD}.get(tone, B.INK2)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=NC)
    c = ws.cell(r, 2, text)
    c.font = Font(name=B.FONT, size=9, bold=tone is not None, color=col)
    c.alignment = W.WRAP
    c.fill = fl
    c.border = W.BOX
    for j in range(3, NC + 1):
        cc = ws.cell(r, j)
        cc.border = W.BOX
        cc.fill = fl
    ws.row_dimensions[r].height = h
    return r + 1


def head(ws, r, labels):
    for j, t in enumerate(labels, 2):
        txt(ws, r, j, t, font=W.HEAD, fill=HEADF)
    ws.row_dimensions[r].height = 30
    return r + 1


def series(ws, r0, d, keys, *, fmt, scale=1.0, act=12, extra=None):
    """اثنا عشر شهراً ثم صفُّ السنة — والنسبتان صيغتان حيّتان"""
    r = r0
    for k in range(12):
        fc = k >= act
        txt(ws, r, 2, M[k] + (" ▸" if fc else ""),
            font=W.BOLD if not fc else Font(name=B.FONT, size=10, color=B.D_GOLD),
            align=W.RGT, fill=FCF if fc else None)
        for j, key in enumerate(keys, 3):
            v = d[key][k] / scale if key else None
            txt(ws, r, j, v, fmt=fmt, fill=FCF if (fc and j == 5) else None)
        if extra:
            extra(ws, r, k)
        r += 1
    # صفُّ السنة
    txt(ws, r, 2, "السنة", font=W.BOLD, align=W.RGT, fill=BANDF)
    for j, key in enumerate(keys, 3):
        col = get_column_letter(j)
        txt(ws, r, j, f"=SUM({col}{r0}:{col}{r-1})", fmt=fmt,
            font=W.BOLD, fill=BANDF)
    if extra:
        extra(ws, r, None)
    ws.row_dimensions[r].height = 20
    return r + 1


def build():
    d = json.load(open(SRC, encoding="utf-8"))
    ra, rr = d["fuel_actual_months"], d["rent_actual_months"]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "مقارنة"
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 26
    for j in range(3, NC + 1):
        ws.column_dimensions[get_column_letter(j)].width = 15

    r = 2
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=NC)
    c = ws.cell(r, 2, "المخطط · الفعلي · الموازنة")
    c.font = W.TITLE
    c.alignment = W.RGT
    ws.row_dimensions[r].height = 28
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=NC)
    c = ws.cell(r, 2, f"المصدر: {d['source']} — "
                      f"الوقود فعليٌّ تسعةَ أشهر، والإيجار ثمانية. "
                      f"والشهر الموسوم ▸ غيرُ فعليّ.")
    c.font = W.SUB
    c.alignment = W.RGT
    r += 2

    # ═════════════ ⓪ الملخّص — ثلاثة أبعادٍ في جدولٍ واحد
    r = band(ws, r, "⓪  الملخّص · المحطات والوقود والوحدات")
    r = head(ws, r, ["البند", "المخطط", "الفعلي", "الموازنة ٢٧",
                     "الفعلي ÷ المخطط", "٢٧ ÷ الفعلي", "المصدر", ""])
    D26 = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

    def lpd(lit, cnt, upto=12):
        sd = sum(cnt[k] * D26[k] for k in range(upto))
        return sum(lit[:upto]) / sd if sd else 0

    def per(blk, k):
        return d[blk][k] / d[blk + "_n"][k] if d[blk + "_n"][k] else 0

    rows = [
        ("محطةٌ قائمةٌ بنهاية السنة",
         d["register_existing"] + d["register_opens_26"],
         d["fuel_l26_n"][ra - 1], d["fuel_has_27"], NUM,
         "السجل مقابل ما باع في سبتمبر"),
        ("محطةٌ باعت في أي شهر", None, d["fuel_sold_actual"],
         d["fuel_has_27"], NUM, "لترٌ > صفر"),
        ("الوقود · مليون لتر", None, sum(d["fuel_l26"][:ra]) / 1e6,
         sum(d["fuel_l27"]) / 1e6, NUM1, "فعليُّ تسعةِ أشهر"),
        ("الوقود · مليون ريال", None, sum(d["fuel_v26"][:ra]) / 1e6,
         sum(d["fuel_v27"]) / 1e6, NUM1, "قبل الضريبة"),
        ("لتر / محطة / يوم", None, lpd(d["fuel_l26"], d["fuel_l26_n"], ra),
         lpd(d["fuel_l27"], d["fuel_l27_n"]), NUM,
         "إنتاجيةُ المحطة الواحدة"),
        ("وحدةٌ تفوتر · مخطط «أ»", d["rent_p26a_n"][11],
         d["rent_gross_n"][rr - 1], d["rent_b27_n"][11], NUM, "ديسمبر"),
        ("وحدةٌ تفوتر · مخطط «ب»", d["rent_p26b_n"][11],
         d["rent_gross_n"][rr - 1], d["rent_b27_n"][11], NUM, "ديسمبر"),
        ("الإيجار · مليون ريال «أ»", sum(d["rent_p26a"]) / 1e6,
         sum(d["rent_gross"]) / 1e6, sum(d["rent_b27"]) / 1e6, NUM1,
         "السنة ▸ بالمنقول"),
        ("الإيجار · مليون ريال «ب»", sum(d["rent_p26b"]) / 1e6,
         sum(d["rent_gross"]) / 1e6, sum(d["rent_b27"]) / 1e6, NUM1,
         "السنة ▸ بالمنقول"),
        ("ريال / وحدة / شهر", per("rent_p26b", 11),
         per("rent_gross", rr - 1), per("rent_b27", 11), NUM,
         "المخطط «ب» · والفعلي أغسطس"),
    ]
    for lbl, p, a, b, fmt, src in rows:
        txt(ws, r, 2, lbl, font=W.BOLD, align=W.RGT)
        txt(ws, r, 3, p if p is not None else "—", fmt=fmt if p else None)
        txt(ws, r, 4, a, fmt=fmt)
        txt(ws, r, 5, b, fmt=fmt)
        txt(ws, r, 6, f'=IF(C{r}="—","—",D{r}/C{r})' if p is not None else "—",
            fmt=PCT, fill=CALCF)
        txt(ws, r, 7, f'=IF(D{r}=0,"—",E{r}/D{r})', fmt=NUM1, fill=CALCF)
        ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=NC)
        c = ws.cell(r, 8, src)
        c.font = W.SMALL
        c.alignment = W.RGT
        c.border = W.BOX
        ws.cell(r, 9).border = W.BOX
        r += 1
    r += 1
    r = note(ws, r, "ومن أين تأتي الموازنة؟  الوقود "
                    f"{d['fuel_has_27']/d['fuel_l26_n'][ra-1]:.2f}× محطاتٍ × "
                    f"{lpd(d['fuel_l27'],d['fuel_l27_n'])/lpd(d['fuel_l26'],d['fuel_l26_n'],ra)*100:.0f}٪ "
                    "إنتاجيةً للمحطة. والإيجار "
                    f"{d['rent_b27_n'][11]/d['rent_gross_n'][rr-1]:.2f}× وحداتٍ × "
                    f"{per('rent_b27',11)/per('rent_gross',rr-1)*100:.0f}٪ "
                    "إيجاراً للوحدة. فالوقود يَعِد توسّعاً وزيادةً يسيرةً في "
                    "الإنتاجية، والإيجار يَعِد عدداً وحدَه — وإيرادُ الوحدة "
                    "فيه ينزل إلى أقلَّ من النصف.", tone="gold", h=36)
    r += 1

    # ═════════════ ① الإيجار · الإيراد
    r = band(ws, r, "①  الإيجار · الإيراد المفوتَر (ريال)")
    r = note(ws, r, "والمخطط نسختان في الملف نفسه: «أ» عمودُ "
                    "`P-26-Revenue` ‹١٬٥٨٩ وحدة بديسمبر›، و«ب» عمودُ "
                    "`P-26-Rental Income` ‹٢٨١ وحدة›. وأيَّهما اتُّخِذ "
                    "أساساً تغيّر الحكم كلُّه — فالفرق بينهما ١٣٫٤ مليون ريال.",
             tone="gold", h=32)
    r = head(ws, r, ["الشهر", "مخطط ٢٦ ‹أ›", "مخطط ٢٦ ‹ب›", "الفعلي ٢٦",
                     "الفعلي ÷ أ", "الفعلي ÷ ب", "موازنة ٢٧", "٢٧ ÷ الفعلي"])

    # خريطة الأعمدة: C مخطط أ · D مخطط ب · E الفعلي · F ÷أ · G ÷ب · H موازنة · I ٢٧÷الفعلي
    def ratios(ws, rw, band=False):
        fl = BANDF if band else CALCF
        fn = W.BOLD if band else W.BLACK
        txt(ws, rw, 6, f'=IF(C{rw}=0,"—",E{rw}/C{rw})', fmt=PCT, fill=fl, font=fn)
        txt(ws, rw, 7, f'=IF(D{rw}=0,"—",E{rw}/D{rw})', fmt=PCT, fill=fl, font=fn)
        txt(ws, rw, 9, f'=IF(E{rw}=0,"—",H{rw}/E{rw})', fmt=NUM1, fill=fl, font=fn)

    def rent_row(ws, rw, k):
        if k is None:
            txt(ws, rw, 8, f"=SUM(H{rw-12}:H{rw-1})", fmt=MONEY,
                font=W.BOLD, fill=BANDF)
        else:
            txt(ws, rw, 8, d["rent_b27"][k], fmt=MONEY)
        ratios(ws, rw, band=k is None)

    r1 = r
    r = series(ws, r, d, ["rent_p26a", "rent_p26b", "rent_gross"],
               fmt=MONEY, act=rr, extra=rent_row)
    # صفُّ الفعلي المعتمد — ثمانيةُ أشهرٍ لا اثنا عشر
    txt(ws, r, 2, f"الفعلي المعتمد ‹ينا–{M[rr-1][:3]}›", font=W.BOLD,
        align=W.RGT, fill=BANDF)
    for j, k in ((3, "C"), (4, "D"), (5, "E")):
        txt(ws, r, j, f"=SUM({k}{r1}:{k}{r1+rr-1})", fmt=MONEY,
            font=W.BOLD, fill=BANDF)
    txt(ws, r, 6, f"=E{r}/C{r}", fmt=PCT, font=W.BOLD, fill=BANDF)
    txt(ws, r, 7, f"=E{r}/D{r}", fmt=PCT, font=W.BOLD, fill=BANDF)
    for j in (8, 9):
        txt(ws, r, j, None, fill=BANDF)
    r += 2
    r = note(ws, r, "على المخطط «أ»: وحداتٌ ٢٨٪ وإيرادٌ ٩٠٪ — لأن الوحدات "
                    "الغائبة مُسعَّرةٌ بـ٤٢٪ من إيجار الوحدة القائمة. وعلى "
                    "«ب»: وحداتٌ ١٤٦٪ وإيرادٌ ٩٩٪ — أجّرنا وحداتٍ أكثر بـ٦٥٪ "
                    "من الإيجار المخطط. فالفجوة في الحالتين سعرُ الوحدة لا "
                    "عددُها — ومن قرأها عدداً سيُطالب بتأجيرٍ أكثر حيث "
                    "المشكلة في الريال لا في الوحدة.", tone="bad", h=34)
    r += 1

    # ═════════════ ② الإيجار · الوحدات
    r = band(ws, r, "②  الإيجار · عدد الوحدات المفوتَرة")
    r = head(ws, r, ["الشهر", "مخطط ٢٦ ‹أ›", "مخطط ٢٦ ‹ب›", "الفعلي ٢٦",
                     "الفعلي ÷ أ", "الفعلي ÷ ب", "موازنة ٢٧", "٢٧ ÷ الفعلي"])
    r0 = r
    for k in range(12):
        fc = k >= rr
        txt(ws, r, 2, M[k] + (" ▸" if fc else ""),
            font=W.BOLD if not fc else Font(name=B.FONT, size=10, color=B.D_GOLD),
            align=W.RGT, fill=FCF if fc else None)
        txt(ws, r, 3, d["rent_p26a_n"][k], fmt=NUM)
        txt(ws, r, 4, d["rent_p26b_n"][k], fmt=NUM)
        txt(ws, r, 5, d["rent_gross_n"][k], fmt=NUM, fill=FCF if fc else None)
        txt(ws, r, 8, d["rent_b27_n"][k], fmt=NUM)
        ratios(ws, r)
        r += 1
    # آخرُ شهرٍ ÷ يناير — وآخرُ شهرٍ في الفعلي هو أغسطس لا ديسمبر
    last = {3: r - 1, 4: r - 1, 8: r - 1, 5: r0 + rr - 1}
    txt(ws, r, 2, "آخر شهر ÷ يناير", font=W.BOLD, align=W.RGT, fill=BANDF)
    for j in (3, 4, 5, 8):
        col = get_column_letter(j)
        txt(ws, r, j, f'=IF({col}{r0}=0,"—",{col}{last[j]}/{col}{r0})',
            fmt=NUM1, font=W.BOLD, fill=BANDF)
    for j in (6, 7, 9):
        txt(ws, r, j, None, fill=BANDF)
    r += 1
    txt(ws, r, 2, "صافي الزيادة شهرياً", font=W.BOLD, align=W.RGT, fill=BANDF)
    for j in (3, 4, 5, 8):
        col = get_column_letter(j)
        span = last[j] - r0          # عددُ الفواصل بين أول شهرٍ وآخره
        txt(ws, r, j, f"=({col}{last[j]}-{col}{r0})/{span}", fmt=NUM1,
            font=W.BOLD, fill=BANDF)
    for j in (6, 7, 9):
        txt(ws, r, j, None, fill=BANDF)
    r += 2
    r = note(ws, r, f"والموازنة تبدأ يناير ٢٧ بـ{d['rent_b27_n'][0]} وحدة "
                    f"مقابل {d['rent_gross_n'][rr-1]} تفوتر اليوم — أي "
                    f"‎+{d['rent_b27_n'][0]-d['rent_gross_n'][rr-1]} وحدةً "
                    f"قبل أن تبدأ السنة، ثم "
                    f"‎+{(d['rent_b27_n'][11]-d['rent_b27_n'][0])/11:.0f} "
                    f"وحدةً شهرياً طوالها. والمحقَّق في ٢٦: "
                    f"‎+{d['rent_gross_n'][rr-1]-d['rent_gross_n'][0]} وحدةً "
                    f"بين يناير وأغسطس = "
                    f"{(d['rent_gross_n'][rr-1]-d['rent_gross_n'][0])/(rr-1):.1f} "
                    f"وحدة/شهر — أي "
                    f"{((d['rent_b27_n'][11]-d['rent_b27_n'][0])/11)/((d['rent_gross_n'][rr-1]-d['rent_gross_n'][0])/(rr-1)):.1f}× "
                    f"الوتيرة المُثبتة.", tone="bad", h=32)
    r = note(ws, r, f"وكلُّ الوحدات الشاغرة في موازنة ٢٧ "
                    f"({d['rent_b27_vacant_notactive']} وحدة) في محطاتٍ "
                    f"`Not Active` لم تُفتح. والقابل للتأجير اليوم "
                    f"({d['rent_pool_active_vacant']} وحدةً شاغرةً في محطاتٍ "
                    f"نشطة) موازنتُه صفر. فنصفُ النمو جدولُ افتتاحاتٍ لا "
                    f"مستهدفٌ تأجيريّ.", tone="gold", h=32)
    r += 1

    # ═════════════ ③ الوقود
    r = band(ws, r, "③  الوقود · اللترات والمحطات")
    r = note(ws, r, "ولا مخطط ٢٠٢٦ في الملف: عمودا `P-2026` موسومان "
                    "«old» وقيمتُهما صفر. فالمقارنة هنا فعليٌّ مقابل موازنة، "
                    "والمخططُ ناقصٌ يُطلَب من المالية.", tone="gold")
    r = head(ws, r, ["الشهر", "فعلي ٢٦ · لتر", "محطة بائعة", "ريال/لتر",
                     "موازنة ٢٧ · لتر", "محطة", "ريال/لتر", "٢٧ ÷ ٢٦"])
    r0 = r
    for k in range(12):
        fc = k >= ra
        txt(ws, r, 2, M[k] + (" ▸" if fc else ""),
            font=W.BOLD if not fc else Font(name=B.FONT, size=10, color=B.D_GOLD),
            align=W.RGT, fill=FCF if fc else None)
        txt(ws, r, 3, d["fuel_l26"][k], fmt=NUM, fill=FCF if fc else None)
        txt(ws, r, 4, d["fuel_l26_n"][k], fmt=NUM, fill=FCF if fc else None)
        txt(ws, r, 5, d["fuel_v26"][k], fmt=None)
        ws.cell(r, 5).value = (f'=IF(C{r}=0,"—",{d["fuel_v26"][k]}/C{r})')
        ws.cell(r, 5).number_format = NUM3
        ws.cell(r, 5).fill = CALCF
        txt(ws, r, 6, d["fuel_l27"][k], fmt=NUM)
        txt(ws, r, 7, d["fuel_l27_n"][k], fmt=NUM)
        txt(ws, r, 8, f'=IF(F{r}=0,"—",{d["fuel_v27"][k]}/F{r})',
            fmt=NUM3, fill=CALCF)
        txt(ws, r, 9, f'=IF(C{r}=0,"—",F{r}/C{r})', fmt=NUM1, fill=CALCF)
        r += 1
    txt(ws, r, 2, "السنة", font=W.BOLD, align=W.RGT, fill=BANDF)
    for j in (3, 6):
        col = get_column_letter(j)
        txt(ws, r, j, f"=SUM({col}{r0}:{col}{r-1})", fmt=NUM,
            font=W.BOLD, fill=BANDF)
    txt(ws, r, 4, d["fuel_any_26"], fmt=NUM, font=W.BOLD, fill=BANDF)
    txt(ws, r, 7, d["fuel_has_27"], fmt=NUM, font=W.BOLD, fill=BANDF)
    txt(ws, r, 5, f'={sum(d["fuel_v26"])}/C{r}', fmt=NUM3,
        font=W.BOLD, fill=BANDF)
    txt(ws, r, 8, f'={sum(d["fuel_v27"])}/F{r}', fmt=NUM3,
        font=W.BOLD, fill=BANDF)
    txt(ws, r, 9, f"=F{r}/C{r}", fmt=NUM1, font=W.BOLD, fill=BANDF)
    rfuel = r
    r += 1
    txt(ws, r, 2, f"الفعلي ‹ينا–{M[ra-1][:3]}›", font=W.BOLD,
        align=W.RGT, fill=BANDF)
    txt(ws, r, 3, f"=SUM(C{r0}:C{r0+ra-1})", fmt=NUM, font=W.BOLD, fill=BANDF)
    txt(ws, r, 4, d["fuel_sold_actual"], fmt=NUM, font=W.BOLD, fill=BANDF)
    txt(ws, r, 5, f'={sum(d["fuel_v26"][:ra])}/C{r}', fmt=NUM3,
        font=W.BOLD, fill=BANDF)
    for j in (6, 7, 8, 9):
        txt(ws, r, j, None, fill=BANDF)
    r += 1
    txt(ws, r, 2, "حصةُ الفعلي من ٢٦", font=W.BOLD, align=W.RGT, fill=BANDF)
    txt(ws, r, 3, f"=C{r-1}/C{rfuel}", fmt=PCT, font=W.BOLD, fill=BANDF)
    for j in (4, 5, 6, 7, 8, 9):
        txt(ws, r, j, None, fill=BANDF)
    r += 2
    r = note(ws, r, f"ومحطاتُ ٢٧ الجديدة ({d['fuel_rate_new_27']['n']} محطة) "
                    f"مُقدَّرةٌ بوسيط "
                    f"{d['fuel_rate_new_27']['median']:,.0f} لتر/يوم، "
                    f"وأدناها {d['fuel_rate_new_27']['min']:,.0f} — "
                    f"مقابل وسيط "
                    f"{d['fuel_rate_old_27']['median']:,.0f} للمحطة القائمة "
                    f"في الخطة نفسها. فلا واحدةَ منها دون المئين ٧٥ من "
                    f"الشبكة العاملة. مُدخَلٌ يُفاوَض قبل الاعتماد.",
             tone="bad", h=32)
    r += 1

    # ═════════════ ④ المحطات
    r = band(ws, r, "④  المحطات · السجل مقابل ما باع فعلاً")
    r = head(ws, r, ["البند", "عدد", "المصدر", "", "", "", "", ""])
    rows = [("سجلُّ المحطات كلُّه", d["register_total"], "ورقة 1│Register"),
            ("قائمة ‹Existing›", d["register_existing"], "عمود Status"),
            ("باعت فعلاً ينا–سبت", d["fuel_sold_actual"], "لترٌ > صفر"),
            ("أعلى شهرٍ فعليّ ‹سبتمبر›", d["fuel_l26_n"][ra - 1], "لترٌ > صفر"),
            ("تفتح نوفمبر–ديسمبر ٢٦", d["register_opens_26"], "«New – opens»"),
            ("تفتح خلال ٢٠٢٧", d["register_opens_27"], "«New – opens»"),
            ("لها خطةُ لترٍ في ٢٧", d["fuel_has_27"], "عمود P-2027"),
            ("منها بلا لترٍ في ٢٦", d["fuel_new_27"], "جديدةٌ كلياً")]
    for lbl, v, src in rows:
        txt(ws, r, 2, lbl, font=W.BOLD, align=W.RGT)
        txt(ws, r, 3, v, fmt=NUM)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=NC)
        c = ws.cell(r, 4, src)
        c.font = W.SMALL
        c.alignment = W.RGT
        c.border = W.BOX
        for j in range(5, NC + 1):
            ws.cell(r, j).border = W.BOX
        r += 1
    r = note(ws, r, f"والفارق يُقرأ هكذا: السجلُّ يَعِد "
                    f"{d['register_existing']} محطةً قائمة، وباع منها "
                    f"{d['fuel_sold_actual']} فعلاً في التسعة أشهر — "
                    f"{d['register_existing']-d['fuel_sold_actual']} محطةً "
                    f"«قائمة» بلا لترٍ واحد. تُفرَز قبل أن يُبنى عليها مستهدف.",
             tone="gold")
    r += 1

    # ═════════════ ملاحظات
    r = band(ws, r, "⑤  ما لا يُقرأ من الجدول — ويُطلَب")
    for t in ("مخططُ الوقود لسنة ٢٠٢٦ — غيرُ موجودٍ في الملف (عمودا P-2026 صفر) · المالية",
              "أيُّ نسختي مخطط الإيجار هي المعتمدة — «أ» ١٬٥٨٩ وحدة أم «ب» ٢٨١ · المالية",
              f"سببُ إشعارات أغسطس الدائنة ({d['rent_credit'][7]/1e6:,.2f} مليون ريال في شهر) · المالية",
              f"{d['rent_billing_zero_b27']['units']} وحدةً تفوتر اليوم وموازنةُ ٢٧ لها صفر "
              f"({d['rent_billing_zero_b27']['annual']/1e6:,.2f} مليون ريال/سنة) · التأجير",
              f"{d['rent_vs_standard']['no_standard']} وحدةً مفوتَرةً بلا إيجارٍ معياريٍّ مسجَّل "
              f"({d['rent_vs_standard']['no_standard_annual']/1e6:,.2f} مليون ريال/سنة) · التأجير",
              "جدولُ افتتاح المحطات بتاريخٍ مُلزم — عليه ٤٧٪ من زيادة الوقود · التطوير"):
        r = note(ws, r, "•  " + t, h=22)

    ws.freeze_panes = "B6"
    ws.print_area = f"A1:{get_column_letter(NC)}{r}"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(OUT)
    print(f"✓ {OUT}")
    print(f"  الإيجار: مخطط أ {sum(d['rent_p26a'])/1e6:,.2f}م · "
          f"ب {sum(d['rent_p26b'])/1e6:,.2f}م · "
          f"فعلي {sum(d['rent_gross'][:d['rent_actual_months']])/1e6:,.2f}م · "
          f"موازنة ٢٧ {sum(d['rent_b27'])/1e6:,.2f}م")
    print(f"  الوقود: فعلي {sum(d['fuel_l26'][:d['fuel_actual_months']])/1e6:,.0f}م لتر · "
          f"سنة ٢٦ {sum(d['fuel_l26'])/1e6:,.0f}م · موازنة ٢٧ {sum(d['fuel_l27'])/1e6:,.0f}م")


if __name__ == "__main__":
    build()
