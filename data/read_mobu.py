# -*- coding: utf-8 -*-
"""قراءة ورقة «Mo Bu» — مبيعات المحطة شهراً بشهر ٢٠٢٥ · ٢٠٢٦ · موازنة ٢٠٢٧

   الورقة كتلةٌ لكل محطة: ثلاث سنواتٍ × منتجاتها (٩١ · ٩٥ · ديزل · ٩٨) وصفُّ
   مجموع، ثم صفّا VS%. وفي ذيلها كتلةُ «Grand Total - 18 stations».

   ثلاث مصائد لا تُقرأ الورقة بدونها:

   ① **كتلة الإجمالي في الذيل بلا اسمٍ في عمود B** — فمن تابع آخر محطةٍ
     أسند إليها إجمالَ الورقة كلِّه. والمطابقة هنا صريحة: مجموع المحطات
     الثمانَ عشرةَ = صفُّ الإجمالي في كل شهر، فرقاً صفراً.

   ② **أكتوبر–ديسمبر ٢٠٢٦ تقديرٌ لا فعل** — تقوله الورقة في ذيلها، ويؤكّده
     تساوي الأشهر حرفياً. فالفعليّ يناير–سبتمبر.

   ③ **عمود VS% في الورقة مقسومٌ على السنة الجديدة لا على الأساس** — تقوله
     الورقة صراحةً. فـ«‎−٨١٪» في عرفات الشرائع تغيُّرُه الحقيقي ‎−٤٥٪.
     يُقرأ عمودها على أنه نسبةٌ من الجديد، لا تغيّراً مئوياً.

   يُشغَّل:  python3 data/read_mobu.py <ملف الموازنة.xlsx>
"""
import json
import sys

import openpyxl

OUT = "outlets/data/mobu.json"
SHEET = "Mo Bu"
ACT26 = 9          # آخر شهرٍ فعليٍّ في ٢٠٢٦ — سبتمبر
DAYS25 = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
DAYS26 = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
COL = dict(idx=0, name=1, model=2, typ=3, note=4, code=5, year=6, prod=7, jan=8)


def num(v):
    return float(v) if isinstance(v, (int, float)) else 0.0


def read(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[SHEET]
    stations, cur, year, grand = [], None, None, {}
    for r in ws.iter_rows(min_row=2, max_col=23, values_only=True):
        a, name = r[COL["idx"]], r[COL["name"]]
        if isinstance(a, str) and "Grand Total" in a:
            cur = "GT"
        elif a not in (None, "") and name:
            cur = dict(idx=a, name=str(name).strip(),
                       model=str(r[COL["model"]] or "").strip(),
                       typ=str(r[COL["typ"]] or "").strip(),
                       note=str(r[COL["note"]] or "").strip(),
                       code=str(r[COL["code"]] or "").strip(), d={})
            stations.append(cur)
        if r[COL["year"]] not in (None, ""):
            year = str(r[COL["year"]]).strip()
        prod = r[COL["prod"]]
        if cur is None or prod in (None, ""):
            continue
        vals = [num(r[COL["jan"] + k]) for k in range(12)]
        key = f"{year}|{str(prod).strip()}"
        (grand if cur == "GT" else cur["d"])[key] = vals

    for s in stations:
        for tag, key in (("a25", "Act 2025|Total"), ("a26", "Act 2026|Total"),
                         ("b27", "BU 2027|Total")):
            s[tag] = s["d"].get(key, [0.0] * 12)
        s["live"] = sum(s["a25"]) + sum(s["a26"]) > 0
        s["mature"] = (all(s["a25"][k] > 0 for k in range(ACT26))
                       and all(s["a26"][k] > 0 for k in range(ACT26)))
        s["first25"] = next((k for k in range(12) if s["a25"][k] > 0), None)
        s["first26"] = next((k for k in range(12) if s["a26"][k] > 0), None)
        s["prods"] = sorted({k.split("|")[1] for k in s["d"]
                             if k.split("|")[1] not in ("Total",)})

    # مطابقة: مجموع المحطات = صفُّ الإجمالي في الورقة
    chk = []
    for k in range(12):
        mine = sum(s["a26"][k] for s in stations)
        theirs = grand.get("Act 2026|Total", [0] * 12)[k]
        chk.append(round(mine - theirs, 2))
    return dict(stations=stations, grand=grand, check26=chk,
                act_months_26=ACT26, days25=DAYS25, days26=DAYS26,
                source=path.split("/")[-1])


def main():
    if len(sys.argv) < 2:
        sys.exit("الاستعمال: python3 data/read_mobu.py <ملف الموازنة.xlsx>")
    d = read(sys.argv[1])
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=1)
    live = [s for s in d["stations"] if s["live"]]
    mat = [s for s in live if s["mature"]]
    print(f"✓ {OUT}")
    print(f"  محطات: {len(d['stations'])} · بحجم: {len(live)} · ناضجة: {len(mat)}")
    print(f"  مطابقة الإجمالي ٢٦: {'صفر في كل شهر' if not any(d['check26']) else d['check26']}")


if __name__ == "__main__":
    main()
