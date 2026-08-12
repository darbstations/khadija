# -*- coding: utf-8 -*-
"""مزامنة ورقة نموذج العامل داخل الخطة التجارية مع النسخة المعتمدة"""
# يُشغَّل من جذر المستودع:  PYTHONPATH=data python3 data/sync_commercial_plan.py
from openpyxl import load_workbook
import worker_model as W

P = "docs/commercial-plan.xlsx"
ST = W.stations()
wb = load_workbook(P)
old = wb.sheetnames.index("نموذج العامل")
print("قبل:", " · ".join(wb.sheetnames))

for name in ("نموذج العامل", W.DATA_SHEET, "مطابقة الوردية مع الطلب"):
    if name in wb.sheetnames:
        del wb[name]

end = W.build_data(wb, ST)                       # تُضاف في النهاية
W.build_worker(wb, ST, end, idx=old)
W.build_shift(wb, ST, idx=old + 1)
wb.save(P)
print("بعد:", " · ".join(wb.sheetnames))
