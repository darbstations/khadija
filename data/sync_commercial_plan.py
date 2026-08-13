# -*- coding: utf-8 -*-
"""مزامنة ورقة نموذج العامل داخل الخطة التجارية مع النسخة المعتمدة"""
# يُشغَّل من جذر المستودع:  PYTHONPATH=data python3 data/sync_commercial_plan.py
from openpyxl import load_workbook
import worker_model as W
import competition_model as C

P = "docs/commercial-plan.xlsx"
ST = W.stations()
FIVE = C.five()
wb = load_workbook(P)
old = wb.sheetnames.index("نموذج العامل")
comp = wb.sheetnames.index("المنافسون")
print("قبل:", " · ".join(wb.sheetnames))

for name in ("نموذج العامل", W.DATA_SHEET, "مطابقة الوردية مع الطلب", C.SHEET):
    if name in wb.sheetnames:
        del wb[name]

# السيطرة الميدانية تلي ورقة المنافسين مباشرة — فهي شاهدها الميداني
C.build_control(wb, FIVE, idx=comp + 1)
# نموذج العامل حُذف للتو، فيُثبَّت موضعه بجاره الباقي لا باسمه
old = wb.sheetnames.index("تجربة الشرايع")
end = W.build_data(wb, ST)                  # تُضاف في النهاية
W.build_worker(wb, ST, end, idx=old)
W.build_shift(wb, ST, idx=old + 1)

# ربط ورقة المنافسين بشاهدها الميداني
cs = wb["المنافسون"]
cs.cell(3, 7, "شاهد ميداني — خمس محطات مكة").font = W.HEAD
cs.cell(3, 7).fill = W.FH; cs.cell(3, 7).alignment = W.CTR; cs.cell(3, 7).border = W.BOX
cs.column_dimensions["G"].width = 42
EV = {
    "الدريس": "٤٠٪ من أقرب المنافسين حول محطاتنا — وأقرب منافس في الفردوس (٤١١ م) والشرايع (٤٨٨ م)",
    "ساسكو": "حضور متفرّق وبعيد حول محطاتنا — الضغط أخف مما توحي الحصة الوطنية",
}
for rr in range(4, cs.max_row + 1):
    nm = str(cs.cell(rr, 2).value or "")
    cell = cs.cell(rr, 7, EV.get(nm, ""))
    cell.alignment = W.WRAP; cell.border = W.BOX; cell.font = W.BLACK
    if nm in EV:
        cell.fill = W.WARN

wb.save(P)
print("بعد:", " · ".join(wb.sheetnames))
