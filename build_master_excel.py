# -*- coding: utf-8 -*-
"""يبني ملفاً موحّداً يجمع كل المؤشرات + المستهدفات + المشاريع + الاستراتيجية في قالب واحد بدلّات حقيقية.
   المصدر: platform/app/kpi_data.py (نفس بيانات المنصة)."""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.drawing.image import Image as XLImage

import sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "platform"))
from app import kpi_data as K

MONTHS = ["يناير","فبراير","مارس","أبريل","مايو","يونيو","يوليو","أغسطس","سبتمبر","أكتوبر","نوفمبر","ديسمبر"]
ORANGE="F47A21"; NAVY="58595B"; BLUE="808285"; STEEL="A7A9AC"; GREEN_IN="C6EFCE"
GOLD="FDE3D1"; WHITE="FFFFFF"; LIGHT="EDEEF0"
thin=Side(style="thin", color="D9D9D9"); BORDER=Border(left=thin,right=thin,top=thin,bottom=thin)
FMT={"pct":"0%","int":"#,##0","num1":"0.0","rial":'#,##0 "ر.س"'}
def font(sz=10,b=False,color="222222"): return Font(name="Tajawal",size=sz,bold=b,color=color)
def fill(c): return PatternFill("solid",fgColor=c)
def C(ws,r,c,v=None,*,f=None,fillc=None,al="right",fmt=None,lock=True,wrap=False,border=True):
    cell=ws.cell(r,c)
    if v is not None: cell.value=v
    cell.font=f or font()
    if fillc: cell.fill=fill(fillc)
    cell.alignment=Alignment(horizontal=al,vertical="center",wrap_text=wrap)
    if fmt: cell.number_format=fmt
    cell.protection=Protection(locked=lock)
    if border: cell.border=BORDER
    return cell

wb=Workbook()
LOGO=os.path.join(os.path.dirname(os.path.abspath(__file__)),"platform/app/static/darb_logo.png")

# =========================================================
# 1) ورقة سجل المؤشرات الموحّد (القالب الأساس بالدلّات الحقيقية)
# =========================================================
REG="سجل المؤشرات"
ws=wb.active; ws.title=REG; ws.sheet_view.rightToLeft=True; ws.sheet_view.showGridLines=False
headers=["#","الإدارة","المحور","المؤشر","الوحدة","القطبية","التجميع","الأولوية","الوزن %",
         "المستهدف","نص المستهدف","الركيزة","المشروع","منظور BSC"]+MONTHS+["YTD","نسبة التحقيق","الحالة"]
widths=[4,18,18,40,9,8,8,12,8,12,16,12,22,16]+[9]*12+[12,12,14]
for i,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(i)].width=w
NC=len(headers); last=get_column_letter(NC)
ws.merge_cells(f"A1:{last}1")
C(ws,1,1,"درب · سجل مؤشرات الأداء الموحّد 2026 — كل المؤشرات والمستهدفات والمشاريع والاستراتيجية",
  f=font(14,True,WHITE),fillc=NAVY,al="center",border=False)
ws.row_dimensions[1].height=28
ws.merge_cells(f"A2:{last}2")
C(ws,2,1,"🟩 الأشهر = إدخال  ·  YTD ونسبة التحقيق والحالة دلّات حقيقية تُحسب تلقائياً  ·  قيم التقنية والتسويق مُعبّأة فعلياً",
  f=font(9,False,NAVY),fillc=GOLD,al="center",border=False)
for c,h in enumerate(headers,1): C(ws,3,c,h,f=font(9,True,WHITE),fillc=BLUE,al="center")
ws.row_dimensions[3].height=30
ws.freeze_panes="O4"
dv=DataValidation(type="decimal",operator="greaterThanOrEqual",formula1="0",allow_blank=True); ws.add_data_validation(dv)

# أعمدة: A1 B2 C3 D4 E5 F6 G7 H8 I9 J10 K11 L12 M13 N14 ; أشهر O15..Z26 ; YTD AA27 ach AB28 status AC29
MO=15; AA="AA"; AB="AB"
r=4; seq=1
for name,key,records in K.DEPARTMENTS:
    ws_w=K.kpi_weights(records)
    for i,(axis,nm,unit,pol,agg,tgt,ttxt,fmt,pillar,project) in enumerate(records):
        C(ws,r,1,seq,f=font(9,True),al="center")
        C(ws,r,2,name,f=font(9),al="right")
        C(ws,r,3,axis,f=font(9),al="right",wrap=True)
        C(ws,r,4,nm,f=font(9),al="right",wrap=True)
        C(ws,r,5,unit,al="center")
        C(ws,r,6,pol,al="center")
        C(ws,r,7,agg,al="center")
        C(ws,r,8,K.priority_label(K.priority(nm)),f=font(8),al="center")
        C(ws,r,9,ws_w[i],fmt="0%",al="center")
        C(ws,r,10,tgt if tgt is not None else "",fmt=FMT[fmt] if tgt is not None else None,al="center")
        C(ws,r,11,ttxt,f=font(8,color="666666"),al="center",wrap=True)
        C(ws,r,12,pillar,f=font(9,color=ORANGE),al="center")
        C(ws,r,13,project,f=font(8),al="right",wrap=True)
        C(ws,r,14,K.perspective(nm),f=font(8),al="center")
        # أشهر — إدخال (تعبئة المُحقَّق في يونيو إن وُجد)
        for m in range(12):
            cc=ws.cell(r,MO+m);
            val = float(K.SEED_ACTUALS[nm]) if (m==5 and nm in K.SEED_ACTUALS) else None
            C(ws,r,MO+m,val,fillc=GREEN_IN,fmt=FMT[fmt],al="center",lock=False)
            dv.add(cc)
        rng=f"$O{r}:$Z{r}"
        ws.cell(r,27).value=(f'=IF($G{r}="SUM",IF(COUNT({rng})=0,"",SUM({rng})),'
            f'IF($G{r}="AVG",IFERROR(AVERAGE({rng}),""),'
            f'IF($G{r}="LAST",IFERROR(LOOKUP(2,1/({rng}<>""),{rng}),""),"")))')
        C(ws,r,27,f=font(9,True),fmt=FMT[fmt],al="center")
        ws.cell(r,28).value=(f'=IF($J{r}="","",IF($AA{r}="","",'
            f'IF($J{r}=0,IF($AA{r}<=0,1,0),'
            f'IF($F{r}="↓",IFERROR($J{r}/$AA{r},0),IFERROR($AA{r}/$J{r},0)))))')
        C(ws,r,28,f=font(9,True,NAVY),fmt="0%",al="center")
        ws.cell(r,29).value=f'=IF($AB{r}="","—",IF($AB{r}>=1,"✅ محقق",IF($AB{r}>=0.85,"🟡 قريب","🔴 تحت الهدف")))'
        C(ws,r,29,f=font(9,True),al="center")
        ws.row_dimensions[r].height=24
        r+=1; seq+=1
REG_END=r-1
ws.auto_filter.ref=f"A3:{last}{REG_END}"
ACH_RANGE=f"'{REG}'!$AB$4:$AB${REG_END}"
DEPT_RANGE=f"'{REG}'!$B$4:$B${REG_END}"
PIL_RANGE=f"'{REG}'!$L$4:$L${REG_END}"
PROJ_RANGE=f"'{REG}'!$M$4:$M${REG_END}"
ST_RANGE=f"'{REG}'!$AC$4:$AC${REG_END}"
PERSP_RANGE=f"'{REG}'!$N$4:$N${REG_END}"

# =========================================================
# 2) ورقة الاستراتيجية والركائز + المشاريع (بدلّات أداء حيّة)
# =========================================================
ws=wb.create_sheet("الاستراتيجية", 0); ws.sheet_view.rightToLeft=True; ws.sheet_view.showGridLines=False
for i,w in enumerate([3,26,40,30,14],1): ws.column_dimensions[get_column_letter(i)].width=w
try:
    img=XLImage(LOGO); img.width=150; img.height=int(150/3.83); ws.add_image(img,"B1"); ws.row_dimensions[1].height=46
except Exception: pass
ws.merge_cells("B2:E2"); C(ws,2,2,"درب · الاستراتيجية والركائز 2026 — قطاع المحطات والعقار",f=font(15,True,WHITE),fillc=NAVY,al="center",border=False)
ws.row_dimensions[2].height=26
for c,h in enumerate(["","الركيزة (5 سنوات)","الوصف","مستهدف 2026"],1): C(ws,4,c,h,f=font(10,True,WHITE),fillc=BLUE,al="center")
strat=[("النمو","ركيزة","التوسع في المواقع وزيادة الوصول وتنمية الإيراد.",""),
 ("التوسع في المواقع","","افتتاح وتشغيل محطات جديدة.","التشغيل 85 · الامتياز 150 محطة"),
 ("الامتياز والعقود","","تنمية الإيراد عبر الامتياز والاستثمار.","الامتياز 167 · الاستثمار 190 عقد"),
 ("زيادة وصول العملاء","","رفع المبيعات والتغطية.","718.8M لتر · 13 منطقة"),
 ("الابتكار","ركيزة","مشاريع جديدة وتحول تقني وهوية وتانكي.",""),
 ("التحول التقني وتانكي","","أتمتة 164 محطة · D365 · تطبيق تانكي (التسويق).","أتمتة 90% · تانكي تصاعدي"),
 ("الاستدامة","ركيزة","جودة التشغيل وتجربة العميل والموظفين.",""),
 ("جودة التشغيل وتجربة العميل","","جاهزية وأمن سيبراني ورضا.","جاهزية 99% · CSAT 90%"),
 ("الاستثمار في الموظفين","","التوطين والتدريب والاستبقاء.","6 دورات · استبقاء 90%"),]
r=5
for nm,kind,desc,tgt in strat:
    isr=kind=="ركيزة"
    C(ws,r,2,nm,f=font(11,True,WHITE if isr else NAVY),fillc=BLUE if isr else None,al="right",wrap=True)
    C(ws,r,3,desc,f=font(9,False,WHITE if isr else "222222"),fillc=BLUE if isr else None,al="right",wrap=True)
    C(ws,r,4,tgt,f=font(9,True,WHITE if isr else ORANGE),fillc=BLUE if isr else None,al="right",wrap=True)
    C(ws,r,1,"",fillc=BLUE if isr else None)
    ws.row_dimensions[r].height=24; r+=1
# المشاريع (الخارطة) بدلّات أداء حيّة من السجل
r+=1; C(ws,r,2,"الخارطة التنفيذية — المشاريع",f=font(12,True,WHITE),fillc=ORANGE,al="center"); ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=5); r+=1
for c,h in zip([2,3,4,5],["المشروع","الركيزة","عدد المؤشرات","الأداء (حي)"]): C(ws,r,c,h,f=font(10,True,WHITE),fillc=BLUE,al="center")
r+=1
for p in K.PROJECTS:
    C(ws,r,2,p,f=font(9),al="right",wrap=True)
    C(ws,r,3,K.PROJECT_PILLAR.get(p,""),f=font(9),al="center")
    C(ws,r,4,f'=COUNTIF({PROJ_RANGE},"{p}")',f=font(9,True),al="center")
    C(ws,r,5,f'=IFERROR(AVERAGEIFS({ACH_RANGE},{PROJ_RANGE},"{p}"),"—")',f=font(9,True,NAVY),fmt="0%",al="center")
    r+=1

# =========================================================
# 3) ورقة اللوحة الموجزة (دلّات تجميع حقيقية)
# =========================================================
ws=wb.create_sheet("اللوحة الموجزة", 1); ws.sheet_view.rightToLeft=True; ws.sheet_view.showGridLines=False
for i,w in enumerate([3,28,16,16,16],1): ws.column_dimensions[get_column_letter(i)].width=w
ws.merge_cells("B2:E2"); C(ws,2,2,"اللوحة الموجزة — تجميع حي من سجل المؤشرات",f=font(15,True,WHITE),fillc=NAVY,al="center",border=False); ws.row_dimensions[2].height=28
# بطاقات
C(ws,4,2,"الإنجاز العام",f=font(11,True,WHITE),fillc=BLUE,al="center")
C(ws,5,2,f'=IFERROR(AVERAGE({ACH_RANGE}),0)',f=font(22,True,ORANGE),fmt="0%",al="center")
for i,(lbl,key) in enumerate([("✅ محقق","✅ محقق"),("🟡 قريب","🟡 قريب"),("🔴 تحت الهدف","🔴 تحت الهدف")]):
    c=3+i; C(ws,4,c,lbl,f=font(10,True,WHITE),fillc=STEEL,al="center")
    C(ws,5,c,f'=COUNTIF({ST_RANGE},"{key}")',f=font(18,True,NAVY),al="center")
# حسب الإدارة
hr=7; C(ws,hr,2,"الأداء حسب الإدارة",f=font(12,True,WHITE),fillc=ORANGE,al="center"); ws.merge_cells(start_row=hr,start_column=2,end_row=hr,end_column=4)
for c,h in zip([2,3,4],["الإدارة","الأداء","عدد المؤشرات"]): C(ws,hr+1,c,h,f=font(10,True,WHITE),fillc=STEEL,al="center")
rr=hr+2
for name,key,records in K.DEPARTMENTS:
    C(ws,rr,2,name,f=font(9),al="right")
    C(ws,rr,3,f'=IFERROR(AVERAGEIFS({ACH_RANGE},{DEPT_RANGE},"{name}"),"—")',f=font(9,True,NAVY),fmt="0%",al="center")
    C(ws,rr,4,f'=COUNTIF({DEPT_RANGE},"{name}")',f=font(9),al="center")
    rr+=1
# حسب الركيزة + المنظور (يمين)
pc=6
ws.column_dimensions[get_column_letter(pc)].width=16; ws.column_dimensions[get_column_letter(pc+1)].width=14
C(ws,hr,pc,"الأداء حسب الركيزة",f=font(12,True,WHITE),fillc=ORANGE,al="center"); ws.merge_cells(start_row=hr,start_column=pc,end_row=hr,end_column=pc+1)
for c,h in zip([pc,pc+1],["الركيزة","الأداء"]): C(ws,hr+1,c,h,f=font(10,True,WHITE),fillc=STEEL,al="center")
rr=hr+2
for p in K.PILLARS:
    C(ws,rr,pc,p,f=font(9),al="right"); C(ws,rr,pc+1,f'=IFERROR(AVERAGEIFS({ACH_RANGE},{PIL_RANGE},"{p}"),"—")',f=font(9,True,NAVY),fmt="0%",al="center"); rr+=1
rr+=1
C(ws,rr,pc,"الأداء حسب منظور BSC",f=font(11,True,WHITE),fillc=BLUE,al="center"); ws.merge_cells(start_row=rr,start_column=pc,end_row=rr,end_column=pc+1); rr+=1
for p in ["مالي","العملاء","العمليات الداخلية","التعلّم والنمو"]:
    C(ws,rr,pc,p,f=font(9),al="right"); C(ws,rr,pc+1,f'=IFERROR(AVERAGEIFS({ACH_RANGE},{PERSP_RANGE},"{p}"),"—")',f=font(9,True,NAVY),fmt="0%",al="center"); rr+=1

for s in wb.worksheets: s.sheet_view.rightToLeft=True
out="درب-سجل-المؤشرات-الموحّد-2026.xlsx"
wb.save(out)
print("saved:",out,os.path.getsize(out),"bytes | KPIs:",REG_END-3)
