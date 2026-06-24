# -*- coding: utf-8 -*-
"""ملف إكسل احترافي موحّد: استراتيجية + مستهدفات 2026 + لوحة + جدول مستقل لكل إدارة + سجل مشاريع.
   دلّات حقيقية · RTL · مرتّب. المصدر: platform/app/kpi_data.py"""
import os, sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.drawing.image import Image as XLImage
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "platform"))
from app import kpi_data as K

ORANGE="F47A21"; NAVY="58595B"; BLUE="808285"; STEEL="A7A9AC"; GREEN_IN="C6EFCE"
GOLD="FDE3D1"; WHITE="FFFFFF"; LIGHT="F2F3F5"; RED="C00000"; RED_FILL="F8D7DA"
thin=Side(style="thin", color="D9D9D9"); BORDER=Border(left=thin,right=thin,top=thin,bottom=thin)
FMT={"pct":"0%","int":"#,##0","num1":"0.0","rial":'#,##0 "ر.س"'}
PERSPS=["مالي","العملاء","العمليات الداخلية","التعلّم والنمو"]
def F_(sz=10,b=False,color="222222"): return Font(name="Tajawal",size=sz,bold=b,color=color)
def fill(c): return PatternFill("solid",fgColor=c)
def C(ws,r,c,v=None,*,f=None,fillc=None,al="right",fmt=None,lock=True,wrap=False,border=True):
    cell=ws.cell(r,c)
    if v is not None: cell.value=v
    cell.font=f or F_()
    if fillc: cell.fill=fill(fillc)
    cell.alignment=Alignment(horizontal=al,vertical="center",wrap_text=wrap)
    if fmt: cell.number_format=fmt
    cell.protection=Protection(locked=lock)
    if border: cell.border=BORDER
    return cell
def rtl(ws): ws.sheet_view.rightToLeft=True; ws.sheet_view.showGridLines=False

# بيانات تجريبية واقعية (deterministic) لتعبئة الفراغات — تُستبدل بالأرقام الفعلية لاحقاً
import hashlib
DEMO=True
def eff_actual(nm,pol,tgt):
    if nm in K.SEED_ACTUALS: return K.SEED_ACTUALS[nm]      # رقم حقيقي من الملفات
    if not DEMO or tgt is None or tgt==0: return None        # اترك بلا رقم (أهداف «صفر» أو غير محدّدة)
    frac=(int(hashlib.md5(nm.encode("utf-8")).hexdigest(),16)%1000)/1000.0
    factor=(0.85+frac*0.33) if pol=="↓" else (0.72+frac*0.33)  # تشتيت لإظهار أخضر/أصفر/أحمر
    val=tgt*factor
    return round(val) if tgt>=1000 else round(val,1) if tgt>=10 else round(val,3)

wb=Workbook(); wb.remove(wb.active)
LOGO=os.path.join(os.path.dirname(os.path.abspath(__file__)),"platform/app/static/darb_logo.png")
def logo(ws,anchor="B1",w=150):
    try:
        im=XLImage(LOGO); im.width=w; im.height=int(w/3.83); ws.add_image(im,anchor); ws.row_dimensions[1].height=46
    except Exception: pass

# ============ جداول الإدارات (كل إدارة ورقة مستقلة) ============
# أعمدة: A# B محور C مؤشر D وحدة E قطبية F أولوية G وزن% H مستهدف I نص المستهدف J ركيزة K مشروع L منظور M المُحقَّق N نسبة التحقيق O الحالة
HEAD=["#","المحور","المؤشر","الوحدة","القطبية","الأولوية","الوزن %","المستهدف","نص المستهدف","الركيزة","المشروع","منظور BSC","المُحقَّق","نسبة التحقيق","الحالة"]
WID=[4,16,40,9,7,12,8,12,16,12,22,15,12,12,14]
DEPT_RANGES=[]
def safe_title(name): return name[:31]
for dname,key,records in K.DEPARTMENTS:
    ws=wb.create_sheet(safe_title(dname)); rtl(ws)
    for i,w in enumerate(WID,1): ws.column_dimensions[get_column_letter(i)].width=w
    last=get_column_letter(len(HEAD))
    ws.merge_cells(f"A1:{last}1"); C(ws,1,1,f"إدارة {dname} · مؤشرات الأداء 2026",f=F_(13,True,WHITE),fillc=NAVY,al="center",border=False); ws.row_dimensions[1].height=24
    ws.merge_cells(f"A2:{last}2"); C(ws,2,1,"🟩 عمود «المُحقَّق» أرقام تجريبية — استبدلها بالفعلية · نسبة التحقيق والحالة دلّات تُحسب تلقائياً",f=F_(9,False,NAVY),fillc=GOLD,al="center",border=False)
    for c,h in enumerate(HEAD,1): C(ws,3,c,h,f=F_(9,True,WHITE),fillc=BLUE,al="center")
    ws.row_dimensions[3].height=28
    dv=DataValidation(type="decimal",operator="greaterThanOrEqual",formula1="0",allow_blank=True); ws.add_data_validation(dv)
    weights=K.kpi_weights(records); start=4
    for i,(axis,nm,unit,pol,agg,tgt,ttxt,fmt,pillar,project) in enumerate(records):
        r=start+i
        C(ws,r,1,i+1,f=F_(9,True),al="center")
        C(ws,r,2,axis,f=F_(9),al="right",wrap=True)
        C(ws,r,3,nm,f=F_(9),al="right",wrap=True)
        C(ws,r,4,unit,al="center")
        C(ws,r,5,pol,al="center")
        C(ws,r,6,K.priority_label(K.priority(nm)),f=F_(8),al="center")
        C(ws,r,7,weights[i],fmt="0%",al="center")
        C(ws,r,8,tgt if tgt is not None else "",fmt=FMT[fmt] if tgt is not None else None,al="center")
        C(ws,r,9,ttxt,f=F_(8,color="666666"),al="center",wrap=True)
        C(ws,r,10,pillar,f=F_(9,color=ORANGE),al="center")
        C(ws,r,11,project,f=F_(8),al="right",wrap=True)
        C(ws,r,12,K.perspective(nm),f=F_(8),al="center")
        ach_v=eff_actual(nm,pol,tgt)
        C(ws,r,13,ach_v if ach_v is not None else None,fillc=GREEN_IN,fmt=FMT[fmt],al="center",lock=False)
        dv.add(ws.cell(r,13))
        ws.cell(r,14).value=(f'=IF($H{r}="","",IF($M{r}="","",IF($H{r}=0,IF($M{r}<=0,1,0),'
            f'IF($E{r}="↓",IFERROR($H{r}/$M{r},0),IFERROR($M{r}/$H{r},0)))))')
        C(ws,r,14,f=F_(9,True,NAVY),fmt="0%",al="center")
        ws.cell(r,15).value=f'=IF($N{r}="","—",IF($N{r}>=1,"✅ محقق",IF($N{r}>=0.85,"🟡 قريب","🔴 تحت الهدف")))'
        C(ws,r,15,f=F_(9,True),al="center")
        ws.row_dimensions[r].height=24
    end=start+len(records)-1
    ws.auto_filter.ref=f"A3:{last}{end}"
    ws.freeze_panes="A4"
    DEPT_RANGES.append((dname,safe_title(dname),start,end))

# ============ قاعدة المؤشرات (تجميع — للوحة والاستراتيجية) ============
cons=wb.create_sheet("قاعدة المؤشرات"); rtl(cons)
for i,w in enumerate([4,20,42,16,24,16,12,14,10],1): cons.column_dimensions[get_column_letter(i)].width=w
for c,h in enumerate(["#","الإدارة","المؤشر","الركيزة","المشروع","منظور BSC","نسبة التحقيق","الحالة","الوزن"],1):
    C(cons,1,c,h,f=F_(9,True,WHITE),fillc=BLUE,al="center")
cr=2; seq=1
for dname,sh,s,e in DEPT_RANGES:
    for rr in range(s,e+1):
        C(cons,cr,1,seq,f=F_(8),al="center")
        C(cons,cr,2,dname,f=F_(8),al="right")
        C(cons,cr,3,f"='{sh}'!C{rr}",f=F_(8),al="right")
        C(cons,cr,4,f"='{sh}'!J{rr}",f=F_(8),al="center")
        C(cons,cr,5,f"='{sh}'!K{rr}",f=F_(8),al="right")
        C(cons,cr,6,f"='{sh}'!L{rr}",f=F_(8),al="center")
        C(cons,cr,7,f"='{sh}'!N{rr}",f=F_(8,True,NAVY),fmt="0%",al="center")
        C(cons,cr,8,f"='{sh}'!O{rr}",f=F_(8),al="center")
        C(cons,cr,9,f"='{sh}'!G{rr}",f=F_(8),fmt="0%",al="center")
        cr+=1; seq+=1
CE=cr-1
ACH=f"'قاعدة المؤشرات'!$G$2:$G${CE}"; DEPC=f"'قاعدة المؤشرات'!$B$2:$B${CE}"
PILC=f"'قاعدة المؤشرات'!$D$2:$D${CE}"; PROJC=f"'قاعدة المؤشرات'!$E$2:$E${CE}"
PERC=f"'قاعدة المؤشرات'!$F$2:$F${CE}"; STC=f"'قاعدة المؤشرات'!$H$2:$H${CE}"

# ============ الاستراتيجية والمستهدفات ============
ws=wb.create_sheet("الاستراتيجية والمستهدفات"); rtl(ws)
for i,w in enumerate([3,26,40,30,14],1): ws.column_dimensions[get_column_letter(i)].width=w
logo(ws,"B1")
ws.merge_cells("B2:E2"); C(ws,2,2,"درب · الاستراتيجية ومستهدفات 2026 — قطاع المحطات والعقار",f=F_(15,True,WHITE),fillc=NAVY,al="center",border=False); ws.row_dimensions[2].height=26
for c,h in enumerate(["","الركيزة (5 سنوات) / الهدف","الوصف","مستهدف 2026"],1): C(ws,4,c,h,f=F_(10,True,WHITE),fillc=BLUE,al="center")
strat=[("النمو","ركيزة","التوسع في المواقع وزيادة الوصول وتنمية الإيراد.",""),
 ("التوسع في المواقع","","افتتاح وتشغيل محطات جديدة.","التشغيل 85 · الامتياز 150 محطة"),
 ("الامتياز والعقود","","تنمية الإيراد عبر الامتياز والاستثمار.","الامتياز 167 · الاستثمار 190 عقد"),
 ("زيادة وصول العملاء","","رفع المبيعات والتغطية الجغرافية.","718.8M لتر · 13 منطقة · 59 مدينة"),
 ("الابتكار","ركيزة","مشاريع جديدة وتحول تقني وهوية وتانكي.",""),
 ("التحول التقني","","أتمتة 164 محطة · D365 · جاهزية الأنظمة.","أتمتة 90% · جاهزية 99%"),
 ("ساحات درب وتانكي","","تشغيل ساحات درب ونمو تطبيق تانكي (التسويق).","تأجير ساحة ≥80% · تانكي تصاعدي"),
 ("الاستدامة","ركيزة","جودة التشغيل وتجربة العميل والموظفين.",""),
 ("تجربة العميل وجودة التشغيل","","رضا وأمن سيبراني وجاهزية.","CSAT 90% · جاهزية 99%"),
 ("منظومة التأجير","","الإشغال والتحصيل واستدامة المستأجرين.","إشغال ≥60% · 246 علامة"),
 ("الاستثمار في الموظفين","","التوطين والتدريب والاستبقاء.","6 دورات · استبقاء 90% · سعودة"),]
r=5
for nm,kind,desc,tgt in strat:
    isr=kind=="ركيزة"
    C(ws,r,2,nm,f=F_(11,True,WHITE if isr else NAVY),fillc=BLUE if isr else None,al="right",wrap=True)
    C(ws,r,3,desc,f=F_(9,False,WHITE if isr else "222222"),fillc=BLUE if isr else None,al="right",wrap=True)
    C(ws,r,4,tgt,f=F_(9,True,WHITE if isr else ORANGE),fillc=BLUE if isr else None,al="right",wrap=True)
    C(ws,r,1,"",fillc=BLUE if isr else None); ws.row_dimensions[r].height=24; r+=1
r+=1; C(ws,r,2,"الخارطة التنفيذية — المشاريع (أداء حي)",f=F_(12,True,WHITE),fillc=ORANGE,al="center"); ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=5); r+=1
for c,h in zip([2,3,4,5],["المشروع","الركيزة","عدد المؤشرات","الأداء (حي)"]): C(ws,r,c,h,f=F_(10,True,WHITE),fillc=BLUE,al="center")
r+=1
for p in K.PROJECTS:
    C(ws,r,2,p,f=F_(9),al="right",wrap=True); C(ws,r,3,K.PROJECT_PILLAR.get(p,""),f=F_(9),al="center")
    C(ws,r,4,f'=COUNTIF({PROJC},"{p}")',f=F_(9,True),al="center")
    C(ws,r,5,f'=IFERROR(AVERAGEIFS({ACH},{PROJC},"{p}"),"—")',f=F_(9,True,NAVY),fmt="0%",al="center"); r+=1

# ============ اللوحة الموجزة ============
ws=wb.create_sheet("اللوحة الموجزة"); rtl(ws)
for i,w in enumerate([3,28,14,16,16],1): ws.column_dimensions[get_column_letter(i)].width=w
ws.merge_cells("B2:E2"); C(ws,2,2,"اللوحة الموجزة — تجميع حي من جداول الإدارات",f=F_(15,True,WHITE),fillc=NAVY,al="center",border=False); ws.row_dimensions[2].height=26
C(ws,4,2,"الإنجاز العام",f=F_(11,True,WHITE),fillc=BLUE,al="center")
C(ws,5,2,f'=IFERROR(AVERAGE({ACH}),0)',f=F_(22,True,ORANGE),fmt="0%",al="center")
for i,(lbl,key) in enumerate([("✅ محقق","✅ محقق"),("🟡 قريب","🟡 قريب"),("🔴 تحت الهدف","🔴 تحت الهدف")]):
    c=3+i; C(ws,4,c,lbl,f=F_(10,True,WHITE),fillc=STEEL,al="center"); C(ws,5,c,f'=COUNTIF({STC},"{key}")',f=F_(18,True,NAVY),al="center")
hr=7; C(ws,hr,2,"الأداء حسب الإدارة",f=F_(12,True,WHITE),fillc=ORANGE,al="center"); ws.merge_cells(start_row=hr,start_column=2,end_row=hr,end_column=3)
for c,h in zip([2,3],["الإدارة","الأداء"]): C(ws,hr+1,c,h,f=F_(10,True,WHITE),fillc=STEEL,al="center")
rr=hr+2
for dname,sh,s,e in DEPT_RANGES:
    C(ws,rr,2,dname,f=F_(9),al="right"); C(ws,rr,3,f'=IFERROR(AVERAGEIFS({ACH},{DEPC},"{dname}"),"—")',f=F_(9,True,NAVY),fmt="0%",al="center"); rr+=1
# ركيزة + منظور (يمين)
pc=4; C(ws,hr,pc,"حسب الركيزة",f=F_(12,True,WHITE),fillc=ORANGE,al="center"); ws.merge_cells(start_row=hr,start_column=pc,end_row=hr,end_column=pc+1)
for c,h in zip([pc,pc+1],["الركيزة","الأداء"]): C(ws,hr+1,c,h,f=F_(10,True,WHITE),fillc=STEEL,al="center")
rr=hr+2
for p in K.PILLARS:
    C(ws,rr,pc,p,f=F_(9),al="right"); C(ws,rr,pc+1,f'=IFERROR(AVERAGEIFS({ACH},{PILC},"{p}"),"—")',f=F_(9,True,NAVY),fmt="0%",al="center"); rr+=1
rr+=1; C(ws,rr,pc,"حسب منظور BSC",f=F_(11,True,WHITE),fillc=BLUE,al="center"); ws.merge_cells(start_row=rr,start_column=pc,end_row=rr,end_column=pc+1); rr+=1
for p in PERSPS:
    C(ws,rr,pc,p,f=F_(9),al="right"); C(ws,rr,pc+1,f'=IFERROR(AVERAGEIFS({ACH},{PERC},"{p}"),"—")',f=F_(9,True,NAVY),fmt="0%",al="center"); rr+=1

# ============ سجل المشاريع (كل إدارة) ============
ws=wb.create_sheet("سجل المشاريع"); rtl(ws)
for i,w in enumerate([4,20,40,10,12,10,14,30],1): ws.column_dimensions[get_column_letter(i)].width=w
ws.merge_cells("A1:H1"); C(ws,1,1,"سجل مشاريع الشركة — لكل إدارة مشاريعها (نسبة الإنجاز = المنجز ÷ الكمية)",f=F_(14,True,WHITE),fillc=NAVY,al="center",border=False); ws.row_dimensions[1].height=24
for c,h in enumerate(["#","الإدارة","المشروع","الكمية","نوع الكمية","المنجز","نسبة الإنجاز","المؤشر المرتبط"],1): C(ws,2,c,h,f=F_(10,True,WHITE),fillc=BLUE,al="center")
r=3
def proj_block(dept, items, kpi_link):
    global r
    ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=8)
    C(ws,r,1,f"▾ مشاريع {dept}",f=F_(10,True,WHITE),fillc=ORANGE,al="right"); r+=1
    for i,(nm,qty,qtype,done,pct) in enumerate(items,1):
        C(ws,r,1,i,f=F_(9,True),al="center"); C(ws,r,2,dept,f=F_(9),al="center")
        C(ws,r,3,nm,f=F_(9),al="right",wrap=True); C(ws,r,4,qty,al="center"); C(ws,r,5,qtype,al="center")
        C(ws,r,6,done,fillc=GREEN_IN,al="center",lock=False)
        ws.cell(r,7).value=f'=IFERROR(F{r}/D{r},0)'; C(ws,r,7,f=F_(9,True,NAVY),fmt="0%",al="center")
        C(ws,r,8,kpi_link,f=F_(8,color="666666"),al="right",wrap=True); ws.row_dimensions[r].height=22; r+=1
proj_block("إدارة التقنية", getattr(K,"IT_PROJECTS",[]), "نسبة إنجاز محفظة مشاريع التحول الرقمي")
# أماكن لمشاريع بقية الإدارات (تُعبّأ عند توفّرها)
for dname,sh,s,e in DEPT_RANGES:
    if dname=="التقنية الرقمية": continue
    ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=8)
    C(ws,r,1,f"▾ مشاريع {dname}",f=F_(10,True,WHITE),fillc=STEEL,al="right"); r+=1
    for _ in range(3):
        for c in range(1,9): C(ws,r,c,None,fillc=GREEN_IN if c in (3,4,5,6) else None,lock=(c not in (3,4,5,6)))
        ws.cell(r,7).value=f'=IFERROR(F{r}/D{r},0)'; C(ws,r,7,f=F_(9,True,NAVY),fmt="0%",al="center")
        C(ws,r,2,dname,f=F_(9),al="center"); r+=1

# ============ بيانات الباوربي (جدول مسطّح بقيَم حقيقية — جاهز لـ Power BI/التحليل) ============
# Power BI لا يعيد حساب معادلات الإكسل، لذا نوفّر نسخة مسطّحة بقيَم محسوبة مسبقاً.
flat=wb.create_sheet("بيانات الباوربي"); rtl(flat)
FH=["الإدارة","المحور","المؤشر","الوحدة","القطبية","الأولوية","الوزن","المستهدف","المُحقَّق","نسبة التحقيق","الحالة","الركيزة","المشروع","منظور BSC"]
for i,w in enumerate([18,16,42,9,7,11,8,12,12,13,15,12,22,16],1): flat.column_dimensions[get_column_letter(i)].width=w
for c,h in enumerate(FH,1): C(flat,1,c,h,f=F_(10,True,WHITE),fillc=NAVY,al="center")
flat.row_dimensions[1].height=26
def ach_calc(pol,tgt,act):
    if tgt is None or act is None: return None
    if tgt==0: return 1.0 if act<=0 else 0.0
    if act==0: return 0.0
    return (tgt/act) if pol=="↓" else (act/tgt)
def status_txt(a):
    if a is None: return "—"
    if a>=1: return "✅ محقق"
    if a>=0.85: return "🟡 قريب"
    return "🔴 تحت الهدف"
fr=2
for dname,key,records in K.DEPARTMENTS:
    weights=K.kpi_weights(records)
    for i,(axis,nm,unit,pol,agg,tgt,ttxt,fmt,pillar,project) in enumerate(records):
        act=eff_actual(nm,pol,tgt); a=ach_calc(pol,tgt,act)
        C(flat,fr,1,dname,f=F_(9),al="right")
        C(flat,fr,2,axis,f=F_(9),al="right")
        C(flat,fr,3,nm,f=F_(9),al="right")
        C(flat,fr,4,unit,al="center")
        C(flat,fr,5,pol,al="center")
        C(flat,fr,6,K.priority_label(K.priority(nm)),f=F_(8),al="center")
        C(flat,fr,7,weights[i],fmt="0%",al="center")
        C(flat,fr,8,tgt if tgt is not None else None,al="center")
        C(flat,fr,9,act if act is not None else None,al="center")
        C(flat,fr,10,a if a is not None else None,fmt="0%",al="center")
        C(flat,fr,11,status_txt(a),f=F_(9),al="center")
        C(flat,fr,12,pillar,f=F_(9),al="center")
        C(flat,fr,13,project,f=F_(8),al="right")
        C(flat,fr,14,K.perspective(nm),f=F_(8),al="center")
        fr+=1
flat.freeze_panes="A2"; flat.auto_filter.ref=f"A1:N{fr-1}"

# ترتيب الأوراق
order=["الاستراتيجية والمستهدفات","اللوحة الموجزة","بيانات الباوربي"]+[sh for _,sh,_,_ in DEPT_RANGES]+["سجل المشاريع","قاعدة المؤشرات"]
wb._sheets.sort(key=lambda s: order.index(s.title) if s.title in order else 99)
for s in wb.worksheets: rtl(s)
cons.sheet_state="hidden"

out="درب-المؤشرات-والمشاريع-الموحّد-2026.xlsx"
wb.save(out)
print("saved:",out,os.path.getsize(out),"bytes | إدارات:",len(DEPT_RANGES),"| مؤشرات:",CE-1)
