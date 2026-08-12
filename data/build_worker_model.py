# -*- coding: utf-8 -*-
"""نموذج مبيعات العامل — ملف مستقل"""
# يُشغَّل من جذر المستودع:  PYTHONPATH=data python3 data/build_worker_model.py
from openpyxl import Workbook
import worker_model as W

ST = W.stations()
wb = Workbook()
wb.remove(wb.active)
end = W.build_data(wb, ST)
W.build_worker(wb, ST, end, idx=0)
W.build_shift(wb, ST, idx=1)
wb.save("docs/نموذج-العامل.xlsx")
print("saved: docs/نموذج-العامل.xlsx", "|", " · ".join(wb.sheetnames))
