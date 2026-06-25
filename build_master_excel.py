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
# مؤشرات «المعالم» (مهام لمرة واحدة) — تُعلَّم بعلامة مميزة بدل أولوية الأداء
_MILESTONE_KW=["اكتمال المستندات","اكتمال أرشفة","أرشفة الملفات","اعتماد الهيكل","اكتمال الأوصاف",
  "ترسيخ القيم","أتمتة عمليات الموارد","تفعيل الأنظمة الأساسية"]
def is_milestone(nm): return any(k in nm for k in _MILESTONE_KW)
def prio_label(nm): return "🏁 معلم" if is_milestone(nm) else K.priority_label(K.priority(nm))

# ===== حوكمة المؤشرات: توليد بطاقة التعريف لكل مؤشر =====
MGR_TITLE={"exec":"الرئيس التنفيذي","franchise":"مدير الامتياز","operations":"مدير التشغيل",
 "investment":"مدير الاستثمار","realestate":"مدير العقار","digital":"مدير التقنية",
 "hr":"مدير الموارد البشرية","marketing":"مدير التسويق","quality":"مدير الجودة",
 "legal":"المدير القانوني","finance":"المدير المالي"}
def round_weights_pct(weights):
    """تقريب الأوزان إلى نِسب مئوية صحيحة مجموعها 100٪ (largest remainder) بحد أدنى 1٪."""
    raw=[w*100 for w in weights]; floor=[int(x) for x in raw]; rem=[x-int(x) for x in raw]
    for i in range(len(floor)):
        if floor[i]==0: floor[i]=1; rem[i]=0
    diff=100-sum(floor)
    order=sorted(range(len(raw)),key=lambda i:rem[i],reverse=(diff>0))
    i=0
    while diff!=0 and order:
        idx=order[i%len(order)]
        if diff>0: floor[idx]+=1; diff-=1
        elif floor[idx]>1: floor[idx]-=1; diff+=1
        i+=1
    return [f/100 for f in floor]
def kpi_source(nm,axis):
    t=nm+" "+axis
    for kws,src in [
     (["إيراد","ربح","هامش","تحصيل","DSO","سداد","موازنة","تدفق","نقدي","العائد","ROE","مستحق","مورد","تدقيق","ميزاني","دين","ذمم","إقفال"],"نظام المحاسبة / ERP (D365)"),
     (["قوقل","Google","خرائط"],"تقييمات خرائط قوقل"),
     (["رضا","CSAT","eNPS","استبيان"],"استبيان دوري"),
     (["تانكي","الولاء","متابع","وصول","تفاعل","حمل","MAU"],"أنظمة التسويق الرقمي / تانكي"),
     (["تدريب","سعودة","استبقاء","توظيف","موظف","شواغر","دوران","هيكل","أوصاف"],"نظام الموارد البشرية (HRMS)"),
     (["مشروع","تحول رقمي","أتمتة","سيبراني","أنظمة","تكامل","Uptime"],"مكتب إدارة المشاريع (PMO)"),
     (["جودة","ملاحظات","مخاطر","مطابقة","تفتيش","QMS","امتثال"],"نظام إدارة الجودة (QMS)"),
     (["قضايا","مذكرة","مذكرات","استشار","صياغة"],"نظام إدارة العقود/القضايا"),
     (["عقد","عقود","امتياز","استثمار","تأجير","إيجار","إشغال","استحواذ"],"نظام إدارة العقود (CRM)"),
     (["محطة","محطات","مبيعات","لتر","مضخ","صيانة","سلامة","تشجير","وقود","تعبئة"],"نظام تشغيل المحطات"),
    ]:
        if any(k in t for k in kws): return src
    return "سجلات الإدارة"
def kpi_freq(nm,ttxt):
    if any(k in nm for k in ["eNPS","رضا","CSAT","استبيان","تدقيق","مراجعة الوثائق"]): return "ربعي"
    if "سنوي" in (ttxt or "") or any(k in nm for k in ["ROE","العائد على حقوق","نمو الإيرادات السنوي","الموازنة"]): return "ربعي"
    return "شهري"
def kpi_formula(fmt,unit,nm):
    if "نسبة" in nm or fmt=="pct": return "(المتحقّق ÷ الإجمالي) × 100٪"
    if fmt=="rial": return "إجمالي القيمة (ر.س) خلال الفترة"
    if any(u in unit for u in ["يوم","أسبوع","ساعة","دقيقة","سنة"]): return f"متوسط {unit} خلال الفترة"
    if "نجمة" in unit: return "متوسط التقييم (من 5)"
    if "نقطة" in unit: return "٪ المروّجين − ٪ المنتقدين"
    return "إجمالي العدد خلال الفترة"
_OUTCOME=["إيراد","ربح","هامش","العائد","ROE","رضا","CSAT","eNPS","نمو","حصة","إشغال","تحصيل","مبيعات",
 "استبقاء","قيمة","تقييم","شكاو","الوصول","EBITDA","ربحية","جاهزية","Uptime","سلامة","امتثال","التزام",
 "حل الشكاو","دقة","مطابقة","إغلاق","فعالية","استدامة","تجديد","انخفاض","خفض","سعودة","توطين","مخاطر",
 "إنجاز محفظة","معالجة","حوادث"]
_ACTIVITY=["عدد ","زيارات","حمل","دورات","تنفيذ","إنجاز التقارير","اكتمال المستندات","اكتمال الأوصاف",
 "أرشفة","تدريب","إعداد","مذكرة","معارض","اعتماد الهيكل"]
def kpi_type(nm):
    if any(k in nm for k in _ACTIVITY): return "⚙️ نشاط"
    if any(k in nm for k in _OUTCOME): return "🎯 نتيجة"
    return "🔧 تشغيلي"
def kpi_def(nm,axis,pillar):
    return f"يقيس {nm} ضمن محور «{axis}»، ويعكس مساهمة الإدارة في ركيزة {pillar}."
def kpi_owner(key,axis): return MGR_TITLE.get(key,"مدير الإدارة")
def kpi_approver(src,key):
    if key=="exec": return "اللجنة التنفيذية"
    if "المحاسبة" in src: return "الإدارة المالية"
    return MGR_TITLE.get(key,"مدير الإدارة")

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
    weights=round_weights_pct(K.kpi_weights(records)); start=4
    for i,(axis,nm,unit,pol,agg,tgt,ttxt,fmt,pillar,project) in enumerate(records):
        r=start+i
        C(ws,r,1,i+1,f=F_(9,True),al="center")
        C(ws,r,2,axis,f=F_(9),al="right",wrap=True)
        C(ws,r,3,nm,f=F_(9),al="right",wrap=True)
        C(ws,r,4,unit,al="center")
        C(ws,r,5,pol,al="center")
        C(ws,r,6,prio_label(nm),f=F_(8),al="center")
        C(ws,r,7,weights[i],fmt="0%",al="center")
        C(ws,r,8,tgt if tgt is not None else "",fmt=FMT[fmt] if tgt is not None else None,al="center")
        C(ws,r,9,ttxt,f=F_(8,color="666666"),al="center",wrap=True)
        C(ws,r,10,pillar,f=F_(9,color=ORANGE),al="center")
        C(ws,r,11,project,f=F_(8),al="right",wrap=True)
        C(ws,r,12,K.perspective(nm),f=F_(8),al="center")
        ach_v=eff_actual(nm,pol,tgt)
        C(ws,r,13,ach_v if ach_v is not None else None,fillc=GREEN_IN,fmt=FMT[fmt],al="center",lock=False)
        dv.add(ws.cell(r,13))
        # نسبة التحقيق — مع سقف 100% (MIN) لمنع تضخّم المتوسط
        ws.cell(r,14).value=(f'=IF($H{r}="","",IF($M{r}="","",MIN(1,IF($H{r}=0,IF($M{r}<=0,1,0),'
            f'IF($E{r}="↓",IFERROR($H{r}/$M{r},0),IFERROR($M{r}/$H{r},0))))))')
        C(ws,r,14,f=F_(9,True,NAVY),fmt="0%",al="center")
        ws.cell(r,15).value=f'=IF($N{r}="","⏳ بانتظار هدف",IF($N{r}>=1,"✅ محقق",IF($N{r}>=0.85,"🟡 قريب","🔴 تحت الهدف")))'
        C(ws,r,15,f=F_(9,True),al="center")
        ws.row_dimensions[r].height=24
    end=start+len(records)-1
    ws.auto_filter.ref=f"A3:{last}{end}"
    ws.freeze_panes="A4"
    DEPT_RANGES.append((dname,safe_title(dname),start,end))

# ============ قاعدة المؤشرات (تجميع — للوحة والاستراتيجية) ============
cons=wb.create_sheet("قاعدة المؤشرات"); rtl(cons)
for i,w in enumerate([4,20,42,16,24,16,12,14,10],1): cons.column_dimensions[get_column_letter(i)].width=w
for c,h in enumerate(["#","الإدارة","المؤشر","الركيزة","المشروع","منظور BSC","نسبة التحقيق","الحالة","الوزن","و×ن","الوزن المتاح"],1):
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
        # أعمدة الترجيح: و×ن = الوزن×التحقيق (للمؤشرات المُعبّأة فقط) · الوزن المتاح = الوزن إن وُجد رقم
        C(cons,cr,10,f'=IF($G{cr}="",0,$G{cr}*$I{cr})',f=F_(8),al="center")
        C(cons,cr,11,f'=IF($G{cr}="",0,$I{cr})',f=F_(8),al="center")
        cr+=1; seq+=1
CE=cr-1
ACH=f"'قاعدة المؤشرات'!$G$2:$G${CE}"; DEPC=f"'قاعدة المؤشرات'!$B$2:$B${CE}"
PILC=f"'قاعدة المؤشرات'!$D$2:$D${CE}"; PROJC=f"'قاعدة المؤشرات'!$E$2:$E${CE}"
PERC=f"'قاعدة المؤشرات'!$F$2:$F${CE}"; STC=f"'قاعدة المؤشرات'!$H$2:$H${CE}"
WXN=f"'قاعدة المؤشرات'!$J$2:$J${CE}"; WAV=f"'قاعدة المؤشرات'!$K$2:$K${CE}"  # و×ن · الوزن المتاح

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

# خريطة وزن كل إدارة (ميزان الإدارات)
DW_BY_NAME={name:K.DEPT_WEIGHTS.get(key,0) for name,key,_ in K.DEPARTMENTS}
EXEC_NAME="الإدارة التنفيذية"

# ============ اللوحة الموجزة ============
ws=wb.create_sheet("اللوحة الموجزة"); rtl(ws)
for i,w in enumerate([3,28,14,16,16],1): ws.column_dimensions[get_column_letter(i)].width=w
for col in ("H","I","J"): ws.column_dimensions[col].hidden=True   # أعمدة مساعدة للترجيح
ws.merge_cells("B2:E2"); C(ws,2,2,"اللوحة الموجزة — إنجاز موزون (وزن المؤشر × وزن الإدارة) · سقف 100%",f=F_(13,True,WHITE),fillc=NAVY,al="center",border=False); ws.row_dimensions[2].height=26
C(ws,4,2,"الإنجاز العام (موزون)",f=F_(11,True,WHITE),fillc=BLUE,al="center")
hr=7
drng=f"H{hr+2}:H{hr+1+len(DEPT_RANGES)}"; wrng=f"I{hr+2}:I{hr+1+len(DEPT_RANGES)}"; hasrng=f"J{hr+2}:J{hr+1+len(DEPT_RANGES)}"
C(ws,5,2,f'=IFERROR(SUMPRODUCT({wrng},{drng})/SUMPRODUCT({wrng},{hasrng}),0)',f=F_(22,True,ORANGE),fmt="0%",al="center")
for i,(lbl,key) in enumerate([("✅ محقق","✅ محقق"),("🟡 قريب","🟡 قريب"),("🔴 تحت الهدف","🔴 تحت الهدف")]):
    c=3+i; C(ws,4,c,lbl,f=F_(10,True,WHITE),fillc=STEEL,al="center")
    C(ws,5,c,f'=COUNTIFS({STC},"{key}",{DEPC},"<>{EXEC_NAME}")',f=F_(18,True,NAVY),al="center")  # دون التنفيذية (تفادي الازدواج)
C(ws,hr,2,"الأداء حسب الإدارة (موزون)",f=F_(12,True,WHITE),fillc=ORANGE,al="center"); ws.merge_cells(start_row=hr,start_column=2,end_row=hr,end_column=3)
for c,h in zip([2,3],["الإدارة","الأداء"]): C(ws,hr+1,c,h,f=F_(10,True,WHITE),fillc=STEEL,al="center")
rr=hr+2
for dname,sh,s,e in DEPT_RANGES:
    C(ws,rr,2,dname,f=F_(9),al="right")
    C(ws,rr,3,f'=IFERROR(SUMIF({DEPC},"{dname}",{WXN})/SUMIF({DEPC},"{dname}",{WAV}),"—")',f=F_(9,True,NAVY),fmt="0%",al="center")
    # أعمدة مساعدة مخفيّة: الأداء الرقمي · وزن الإدارة · هل لديها بيانات
    C(ws,rr,8,f'=IFERROR(SUMIF({DEPC},"{dname}",{WXN})/SUMIF({DEPC},"{dname}",{WAV}),0)',f=F_(8),al="center")
    C(ws,rr,9,round(DW_BY_NAME.get(dname,0),5),f=F_(8),al="center")
    C(ws,rr,10,f'=IF(SUMIF({DEPC},"{dname}",{WAV})=0,0,1)',f=F_(8),al="center")
    rr+=1
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
FH=["الإدارة","المحور","المؤشر","النوع","الوحدة","القطبية","الأولوية","الوزن","المستهدف","المُحقَّق","نسبة التحقيق","الحالة","الركيزة","المشروع","منظور BSC","المالك","مصدر البيانات","الدورية"]
for i,w in enumerate([18,15,40,11,8,7,11,8,11,11,12,14,11,20,15,16,22,9],1): flat.column_dimensions[get_column_letter(i)].width=w
for c,h in enumerate(FH,1): C(flat,1,c,h,f=F_(10,True,WHITE),fillc=NAVY,al="center")
flat.row_dimensions[1].height=26
def ach_calc(pol,tgt,act):
    if tgt is None or act is None: return None
    if tgt==0: return 1.0 if act<=0 else 0.0
    if act==0: return 0.0
    return min(1.0,(tgt/act) if pol=="↓" else (act/tgt))   # سقف 100%
def status_txt(a):
    if a is None: return "⏳ بانتظار هدف"
    if a>=1: return "✅ محقق"
    if a>=0.85: return "🟡 قريب"
    return "🔴 تحت الهدف"
fr=2
for dname,key,records in K.DEPARTMENTS:
    weights=round_weights_pct(K.kpi_weights(records))
    for i,(axis,nm,unit,pol,agg,tgt,ttxt,fmt,pillar,project) in enumerate(records):
        act=eff_actual(nm,pol,tgt); a=ach_calc(pol,tgt,act); src=kpi_source(nm,axis)
        C(flat,fr,1,dname,f=F_(9),al="right")
        C(flat,fr,2,axis,f=F_(9),al="right")
        C(flat,fr,3,nm,f=F_(9),al="right")
        C(flat,fr,4,kpi_type(nm),f=F_(8),al="center")
        C(flat,fr,5,unit,al="center")
        C(flat,fr,6,pol,al="center")
        C(flat,fr,7,prio_label(nm),f=F_(8),al="center")
        C(flat,fr,8,weights[i],fmt="0%",al="center")
        C(flat,fr,9,tgt if tgt is not None else None,al="center")
        C(flat,fr,10,act if act is not None else None,al="center")
        C(flat,fr,11,a if a is not None else None,fmt="0%",al="center")
        C(flat,fr,12,status_txt(a),f=F_(9),al="center")
        C(flat,fr,13,pillar,f=F_(9),al="center")
        C(flat,fr,14,project,f=F_(8),al="right")
        C(flat,fr,15,K.perspective(nm),f=F_(8),al="center")
        C(flat,fr,16,kpi_owner(key,axis),f=F_(8),al="center")
        C(flat,fr,17,src,f=F_(8),al="right")
        C(flat,fr,18,kpi_freq(nm,ttxt),f=F_(8),al="center")
        fr+=1
flat.freeze_panes="D2"; flat.auto_filter.ref=f"A1:R{fr-1}"

# ============ بطاقة تعريف المؤشرات (قاموس البيانات — حوكمة) ============
dic=wb.create_sheet("بطاقة تعريف المؤشرات"); rtl(dic)
DH=["#","الإدارة","المؤشر","التعريف","طريقة الاحتساب","النوع","الهدف الاستراتيجي (ركيزة · مشروع)","منظور BSC","الوزن","المالك","مصدر البيانات","الدورية","جهة اعتماد البيانات","آخر تحديث للبيانات"]
for i,w in enumerate([4,16,38,46,30,11,30,16,8,16,24,9,18,16],1): dic.column_dimensions[get_column_letter(i)].width=w
dic.merge_cells(f"A1:{get_column_letter(len(DH))}1"); C(dic,1,1,"بطاقة تعريف المؤشرات — قاموس البيانات والحوكمة (مالك · مصدر · دورية · تعريف · احتساب · اعتماد)",f=F_(12,True,WHITE),fillc=NAVY,al="center",border=False); dic.row_dimensions[1].height=24
for c,h in enumerate(DH,1): C(dic,2,c,h,f=F_(9,True,WHITE),fillc=BLUE,al="center")
dic.row_dimensions[2].height=26
dr=3; seq=1
for dname,key,records in K.DEPARTMENTS:
    weights=round_weights_pct(K.kpi_weights(records))
    for i,(axis,nm,unit,pol,agg,tgt,ttxt,fmt,pillar,project) in enumerate(records):
        src=kpi_source(nm,axis)
        C(dic,dr,1,seq,f=F_(8),al="center")
        C(dic,dr,2,dname,f=F_(8),al="right")
        C(dic,dr,3,nm,f=F_(9),al="right",wrap=True)
        C(dic,dr,4,kpi_def(nm,axis,pillar),f=F_(8,color="444444"),al="right",wrap=True)
        C(dic,dr,5,kpi_formula(fmt,unit,nm),f=F_(8),al="right",wrap=True)
        C(dic,dr,6,kpi_type(nm),f=F_(8),al="center")
        C(dic,dr,7,f"{pillar} · {project}",f=F_(8),al="right",wrap=True)
        C(dic,dr,8,K.perspective(nm),f=F_(8),al="center")
        C(dic,dr,9,weights[i],fmt="0%",al="center")
        C(dic,dr,10,kpi_owner(key,axis),f=F_(8),al="center")
        C(dic,dr,11,src,f=F_(8),al="right",wrap=True)
        C(dic,dr,12,kpi_freq(nm,ttxt),f=F_(8),al="center")
        C(dic,dr,13,kpi_approver(src,key),f=F_(8),al="center")
        C(dic,dr,14,None,fillc=GREEN_IN,al="center",lock=False)  # يُدخله مالك المؤشر
        dic.row_dimensions[dr].height=30; dr+=1; seq+=1
dic.freeze_panes="C3"; dic.auto_filter.ref=f"A2:{get_column_letter(len(DH))}{dr-1}"

# ============ مراجعة وحوكمة المؤشرات (تحليل) ============
gov=wb.create_sheet("مراجعة وحوكمة المؤشرات"); rtl(gov)
for i,w in enumerate([4,24,14,14,14,14,16],1): gov.column_dimensions[get_column_letter(i)].width=w
gov.merge_cells("A1:G1"); C(gov,1,1,"مراجعة وحوكمة المؤشرات — التوازن · النوع · الأوزان · تغطية الأهداف",f=F_(13,True,WHITE),fillc=NAVY,al="center",border=False); gov.row_dimensions[1].height=24
# تجميع إحصائي
from collections import Counter,defaultdict
dep_cnt=Counter(); dep_out=Counter(); dep_act=Counter(); dep_ops=Counter(); dep_wsum=defaultdict(float)
persp_cnt=Counter(); proj_cnt=Counter()
for dname,key,records in K.DEPARTMENTS:
    weights=round_weights_pct(K.kpi_weights(records))
    for i,(axis,nm,unit,pol,agg,tgt,ttxt,fmt,pillar,project) in enumerate(records):
        dep_cnt[dname]+=1; dep_wsum[dname]+=weights[i]
        t=kpi_type(nm)
        if "نتيجة" in t: dep_out[dname]+=1
        elif "نشاط" in t: dep_act[dname]+=1
        else: dep_ops[dname]+=1
        persp_cnt[K.perspective(nm)]+=1; proj_cnt[project]+=1
r=3
C(gov,r,1,"أ) لكل إدارة: العدد · مجموع الأوزان · النوع",f=F_(11,True,WHITE),fillc=ORANGE,al="center"); gov.merge_cells(start_row=r,start_column=1,end_row=r,end_column=7); r+=1
for c,h in enumerate(["#","الإدارة","عدد المؤشرات","مجموع الأوزان","🎯 نتيجة","⚙️ نشاط","🔧 تشغيلي"],1): C(gov,r,c,h,f=F_(9,True,WHITE),fillc=BLUE,al="center")
r+=1; s=1
for dname,key,_ in K.DEPARTMENTS:
    C(gov,r,1,s,f=F_(8),al="center"); C(gov,r,2,dname,f=F_(8),al="right")
    C(gov,r,3,dep_cnt[dname],f=F_(8),al="center")
    C(gov,r,4,round(dep_wsum[dname],2),fmt="0%",f=F_(8,True,"3FB27F" if abs(dep_wsum[dname]-1)<0.005 else "C00000"),al="center")
    C(gov,r,5,dep_out[dname],f=F_(8),al="center"); C(gov,r,6,dep_act[dname],f=F_(8),al="center"); C(gov,r,7,dep_ops[dname],f=F_(8),al="center")
    r+=1; s+=1
r+=1; C(gov,r,1,"ب) توزيع منظورات Balanced Scorecard",f=F_(11,True,WHITE),fillc=ORANGE,al="center"); gov.merge_cells(start_row=r,start_column=1,end_row=r,end_column=7); r+=1
for c,h in enumerate(["#","المنظور","عدد المؤشرات","النسبة"],1): C(gov,r,c,h,f=F_(9,True,WHITE),fillc=BLUE,al="center")
r+=1; s=1; tot=sum(persp_cnt.values())
for p in PERSPS:
    C(gov,r,1,s,f=F_(8),al="center"); C(gov,r,2,p,f=F_(8),al="right"); C(gov,r,3,persp_cnt[p],f=F_(8),al="center")
    C(gov,r,4,round(persp_cnt[p]/tot,3),fmt="0%",f=F_(8),al="center"); r+=1; s+=1
r+=1; C(gov,r,1,"ج) تغطية الأهداف الاستراتيجية (عدد المؤشرات لكل مشروع)",f=F_(11,True,WHITE),fillc=ORANGE,al="center"); gov.merge_cells(start_row=r,start_column=1,end_row=r,end_column=7); r+=1
for c,h in enumerate(["#","المشروع/الهدف","الركيزة","عدد المؤشرات","التغطية"],1): C(gov,r,c,h,f=F_(9,True,WHITE),fillc=BLUE,al="center")
r+=1; s=1
for p in K.PROJECTS:
    n=proj_cnt.get(p,0)
    cov="🔴 فجوة" if n==0 else ("🟡 ضعيفة" if n<3 else "🟢 كافية")
    C(gov,r,1,s,f=F_(8),al="center"); C(gov,r,2,p,f=F_(8),al="right"); C(gov,r,3,K.PROJECT_PILLAR.get(p,""),f=F_(8),al="center")
    C(gov,r,4,n,f=F_(8),al="center"); C(gov,r,5,cov,f=F_(8),al="center"); r+=1; s+=1
gov.freeze_panes="A2"

# ترتيب الأوراق
order=["الاستراتيجية والمستهدفات","اللوحة الموجزة","مراجعة وحوكمة المؤشرات","بطاقة تعريف المؤشرات","بيانات الباوربي"]+[sh for _,sh,_,_ in DEPT_RANGES]+["سجل المشاريع","قاعدة المؤشرات"]
wb._sheets.sort(key=lambda s: order.index(s.title) if s.title in order else 99)
for s in wb.worksheets: rtl(s)
cons.sheet_state="hidden"

out="درب-المؤشرات-والمشاريع-الموحّد-2026.xlsx"
wb.save(out)
print("saved:",out,os.path.getsize(out),"bytes | إدارات:",len(DEPT_RANGES),"| مؤشرات:",CE-1)
