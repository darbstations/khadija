# -*- coding: utf-8 -*-
"""يبني ملف مؤشرات أداء درب 2026 — ملف موحّد بمعادلات حقيقية وحماية خلايا."""
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.protection import WorkbookProtection

PWD = "darb2026"
MONTHS = ["يناير","فبراير","مارس","أبريل","مايو","يونيو","يوليو","أغسطس","سبتمبر","أكتوبر","نوفمبر","ديسمبر"]

# ---------- ألوان وأنماط ----------
NAVY   = "1F3864"; BLUE = "2E5496"; STEEL = "8EAADB"
GREEN_IN = "C6EFCE"   # خلايا الإدخال (غير مقفولة)
GREY   = "F2F2F2"; AXIS = "D9E1F2"; GOLD = "FFE699"
WHITE  = "FFFFFF"
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

def font(sz=11, b=False, color="000000"): return Font(name="Arial", size=sz, bold=b, color=color)
def fill(c): return PatternFill("solid", fgColor=c)
def center(wrap=True): return Alignment(horizontal="center", vertical="center", wrap_text=wrap)
def right(wrap=True):  return Alignment(horizontal="right",  vertical="center", wrap_text=wrap)

LOCKED   = Protection(locked=True)
UNLOCKED = Protection(locked=False)

FMT = {"pct":"0%", "pct1":"0.0%", "int":"#,##0", "num1":"0.0", "rial":'#,##0 "ر.س"'}

def style_cell(c, *, value=None, f=None, fillc=None, align=None, fmt=None, lock=True, border=True):
    if value is not None: c.value = value
    c.font = f or font()
    if fillc: c.fill = fill(fillc)
    c.alignment = align or right()
    if fmt: c.number_format = fmt
    c.protection = LOCKED if lock else UNLOCKED
    if border: c.border = BORDER
    return c

def protect(ws):
    ws.protection.sheet = True
    ws.protection.password = PWD
    ws.protection.formatCells = False
    ws.protection.selectLockedCells = True
    ws.protection.selectUnlockedCells = True
    ws.protection.sort = True
    ws.protection.autoFilter = True
    ws.sheet_view.rightToLeft = True

# =======================================================================
wb = Workbook()

# ---------------- 1) دليل الاستخدام ----------------
ws = wb.active; ws.title = "دليل الاستخدام"
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 3
ws.column_dimensions["B"].width = 110
ws.merge_cells("B2:B2")
style_cell(ws["B2"], value="درب · منظومة مؤشرات الأداء 2026 — ملف موحّد", f=font(18,True,NAVY), fillc=None, align=right(False), border=False)
guide = [
 ("", ""),
 ("🧭 الفكرة", "ملف واحد يتدرّج من فوق لتحت: الاستراتيجية ← لوحة الإدارة التنفيذية ← مؤشرات كل إدارة ← تقييم كل موظف. كل المؤشرات تغذّي نفس اللوحة تلقائياً."),
 ("🟩 خلايا الإدخال", "الخلايا الخضراء فقط قابلة للتعديل — يدخل فيها المسؤول أرقامه الشهرية. باقي الخلايا (المؤشر/المستهدف/المعادلة) مقفولة ومحمية."),
 ("🔢 المعادلات", "نسبة التحقيق وYTD والحالة والدرجات تُحسب آلياً بمعادلات حقيقية. لا تُكتب يدوياً."),
 ("🔒 الحماية", f"كل شيت محمي بكلمة مرور: {PWD}  — لفك الحماية مؤقتاً للتعديل على الهيكل: مراجعة ← حماية الورقة ← أدخل كلمة المرور."),
 ("👤 الصلاحيات", "عند التوزيع: كل موظف/مسؤول يُمنح صلاحية الإدخال على نطاق خلاياه الخضراء فقط (Allow Edit Ranges في Excel / صلاحيات الورقة في SharePoint). الهيكل والمستهدفات تبقى مقفولة للجميع."),
 ("📊 الحالة", "✅ محقق = 100%+   |   🟡 قريب = 85%–99%   |   🔴 تحت الهدف = أقل من 85%   |   — = بانتظار إدخال / مستهدف وصفي."),
 ("🏢 الإدارات", "هذا الإصدار يحتوي نموذج إدارة الامتياز التجاري كاملاً. الإدارات الأخرى (التشغيل، الاستثمار، العقار) تُضاف لنفس الملف وتغذّي نفس اللوحة التنفيذية."),
]
r = 3
for a,b in guide:
    style_cell(ws.cell(r,2, (a+"  —  "+b) if a else b), f=font(11, bool(a)), align=right(True), border=False)
    ws.row_dimensions[r].height = 30 if b else 8
    r += 1
protect(ws)

# ---------------- 2) الاستراتيجية ----------------
ws = wb.create_sheet("الاستراتيجية")
ws.sheet_view.showGridLines = False
widths = [3,26,42,30];
for i,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(i)].width = w
ws.merge_cells("B2:D2")
style_cell(ws["B2"], value="ركائز واستراتيجية درب 2026", f=font(18,True,WHITE), fillc=NAVY, align=center(False), border=False)
ws.row_dimensions[2].height = 32
hdr = ["","الركيزة / الهدف","الوصف","المستهدف الرقمي 2026"]
for c,h in enumerate(hdr,1):
    style_cell(ws.cell(4,c,h), f=font(11,True,WHITE), fillc=BLUE, align=center())
strat = [
 ("النمو", "ركيزة", "التوسع في المواقع وزيادة الوصول للعملاء وتنمية الإيرادات.", ""),
 ("التوسع في المواقع","","افتتاح وتشغيل محطات جديدة.","التشغيل: 85 محطة · الامتياز: 150 محطة"),
 ("الامتياز والعقود","","تنمية الإيرادات عبر الامتياز والاستثمار.","الامتياز: 167 عقد · الاستثمار: 190 عقد"),
 ("زيادة وصول العملاء","","رفع المبيعات والتغطية الجغرافية.","718.8M لتر · 13 منطقة · 59 مدينة"),
 ("الابتكار","ركيزة","مشاريع جديدة وتحول تقني وتطوير الهوية.",""),
 ("الهوية والتصاميم","","تطبيق الهوية المؤسسية في المحطات.","التزام 55% بالهوية"),
 ("التحول التقني وساحات درب","","أتمتة المحطات وتطوير ساحات درب وتطبيق تانكي.","أتمتة 90% · تأجير ساحة درب ≥80%"),
 ("الاستدامة","ركيزة","جودة التشغيل وتجربة العميل والاستثمار في الموظفين.",""),
 ("تجربة العملاء","","رفع الرضا وجودة التجربة.","قوقل ≥4.5–4.7 · CSAT ≥90%"),
 ("جودة التشغيل","","كفاءة وجودة عمليات المحطات.","جاهزية 99% · سلامة 100%"),
 ("منظومة التأجير (العقار)","","الإشغال والتحصيل واستدامة المستأجرين.","إشغال ≥60% · 246 علامة · تحصيل 95%+"),
 ("الاستثمار في الموظفين","","التطوير والاستبقاء.","استبقاء الشركاء 90%+"),
]
r=5
for name,kind,desc,tgt in strat:
    isr = kind=="ركيزة"
    style_cell(ws.cell(r,2,name), f=font(12,True, WHITE if isr else NAVY), fillc=BLUE if isr else None, align=right(True))
    style_cell(ws.cell(r,3,desc), f=font(11,False, WHITE if isr else "000000"), fillc=BLUE if isr else None, align=right(True))
    style_cell(ws.cell(r,4,tgt),  f=font(11,True, WHITE if isr else BLUE), fillc=BLUE if isr else None, align=right(True))
    style_cell(ws.cell(r,1,""), fillc=BLUE if isr else None)
    ws.row_dimensions[r].height = 26
    r+=1
protect(ws)

# =======================================================================
# بيانات مؤشرات إدارة الامتياز (36 مؤشر)
AX = ["النمو والتوسع","المبيعات والأداء التجاري","الأداء المالي","التشغيل والجودة",
      "الهوية والتسويق","التوسع الإقليمي","تجربة العملاء والشركاء","استدامة الشراكة والامتثال"]
# (المحور, الاسم, الوحدة, القطبية, التجميع, المستهدف, نص المستهدف, التنسيق)
KPIS = [
 (0,"عدد محطات الامتياز المشغّلة (تراكمي)","محطة","↑","LAST",150,"150 محطة","int"),
 (0,"عدد عقود الامتياز الجديدة الموقّعة","عقد","↑","SUM",167,"167 عقد","int"),
 (1,"إجمالي مبيعات وقود محطات الامتياز","لتر","↑","SUM",718838077,"718.8M لتر","int"),
 (1,"نسبة المحطات التي حققت مستهدف المبيعات","%","↑","AVG",0.90,"90%+","pct"),
 (1,"متوسط مبيعات المحطة الواحدة شهرياً","لتر/محطة","↑","AVG",None,"+10% نمو شهري","int"),
 (1,"نسبة تأجير وحدات الامتياز","%","↑","AVG",0.70,"70%","pct"),
 (2,"إجمالي إيرادات إدارة الامتياز","ر.س","↑","SUM",None,"تصاعدي مستمر","rial"),
 (2,"رفع هامش الربحية للامتياز","هللة/لتر","↑","AVG",1.8,"+1.8 هللة/لتر","num1"),
 (2,"نسبة التحصيل المالي من شركاء الامتياز","%","↑","AVG",0.95,"95%+","pct"),
 (2,"متوسط أيام التحصيل (DSO)","يوم","↓","AVG",30,"≤ 30 يوم","num1"),
 (2,"إجمالي المستحقات المتأخرة (+30 يوم)","ر.س","↓","LAST",None,"تراجع مستمر","rial"),
 (2,"نسبة الشركاء الملتزمين بمواعيد السداد","%","↑","AVG",0.90,"90%+","pct"),
 (2,"معدل دوران رأس المال (أيام الدورة)","يوم","↓","AVG",5,"≤ 5 أيام","num1"),
 (3,"نسبة جاهزية محطات الامتياز (صفر انقطاع)","%","↑","AVG",0.99,"99%+","pct"),
 (3,"عدد حالات انقطاع الوقود الموثّقة","حالة","↓","SUM",0,"0 حالة","int"),
 (3,"نسبة إغلاق ملاحظات الجودة الميدانية","%","↑","AVG",0.95,"95%+","pct"),
 (3,"متوسط درجة تقييم جودة المحطة (قائمة التحقق)","درجة/100","↑","AVG",85,"85+","num1"),
 (3,"نسبة تنفيذ الزيارات الميدانية المجدولة","%","↑","AVG",1.0,"100%","pct"),
 (3,"نسبة المحطات المستوفية لمعايير السلامة","%","↑","AVG",1.0,"100%","pct"),
 (3,"متوسط وقت معالجة البلاغات التشغيلية","ساعة","↓","AVG",24,"≤ 24 ساعة","num1"),
 (3,"نسبة أتمتة محطات الامتياز","%","↑","LAST",0.90,"90%","pct"),
 (3,"نسبة تطبيق المعيار الشامل داخل المحطات","%","↑","LAST",0.30,"30%","pct"),
 (4,"نسبة التزام المحطات بالهوية المؤسسية","%","↑","LAST",0.55,"55%","pct"),
 (5,"نسبة العقود التي تغطي المناطق الـ13","%","↑","LAST",1.0,"100%","pct"),
 (5,"نسبة العقود في الـ12 مدينة الرئيسية","%","↓","LAST",0.40,"< 40%","pct"),
 (5,"نسبة العقود خارج المدن الرئيسية","%","↑","LAST",0.60,"> 60%","pct"),
 (5,"نسبة العقود على الطرق الإقليمية","%","↑","LAST",0.25,"> 25%","pct"),
 (6,"متوسط تقييم المحطات على خرائط قوقل","نجمة/5","↑","AVG",4.7,"≥ 4.7 نجمة","num1"),
 (6,"معدل رضا شركاء الامتياز (CSAT)","%","↑","AVG",0.90,"90%+","pct"),
 (6,"نسبة حل الشكاوى خلال المدة المحددة","%","↑","AVG",0.95,"95%+","pct"),
 (6,"نسبة انخفاض الشكاوى التشغيلية والإدارية","%","↑","AVG",None,"-30%","pct"),
 (7,"نسبة استدامة أصحاب الامتياز (Retention)","%","↑","LAST",0.90,"90%","pct"),
 (7,"نسبة زيادة عدد أصحاب الامتياز","%","↑","LAST",1.0,"100% نمو","pct"),
 (7,"نسبة اكتمال المستندات النظامية قبل التوقيع","%","↑","AVG",1.0,"100%","pct"),
 (7,"متوسط مدة إغلاق عقد الامتياز","أسبوع","↓","AVG",1,"≤ أسبوع","num1"),
 (7,"نسبة إنجاز التقارير الدورية في موعدها","%","↑","AVG",1.0,"100%","pct"),
]

SH_KPI = "الامتياز · المؤشرات"
# ---------------- 4) مؤشرات الإدارة ----------------
ws = wb.create_sheet(SH_KPI)
ws.sheet_view.showGridLines = False
# الأعمدة: A#  B المحور  C المؤشر  D الوحدة  E القطبية  F التجميع  G المستهدف  H نص المستهدف  I..T الأشهر  U YTD  V نسبة التحقيق  W الحالة
col_w = [4,20,40,10,8,9,12,16] + [9]*12 + [12,12,14]
for i,w in enumerate(col_w,1): ws.column_dimensions[get_column_letter(i)].width = w
ws.merge_cells("A1:W1")
style_cell(ws["A1"], value="إدارة الامتياز التجاري والمبيعات · مؤشرات الأداء 2026  (36 مؤشر / 8 محاور)", f=font(15,True,WHITE), fillc=NAVY, align=center(False), border=False)
ws.row_dimensions[1].height = 30
ws.merge_cells("A2:W2")
style_cell(ws["A2"], value="🟩 الخلايا الخضراء = إدخال شهري للمسؤول   ·   باقي الأعمدة محسوبة/مقفولة   ·   YTD ونسبة التحقيق والحالة تلقائية", f=font(10,False,NAVY), fillc=GOLD, align=center(False), border=False)
header = ["#","المحور","المؤشر","الوحدة","القطبية","التجميع","المستهدف","نص المستهدف"] + MONTHS + ["YTD","نسبة التحقيق","الحالة"]
for c,h in enumerate(header,1):
    style_cell(ws.cell(3,c,h), f=font(10,True,WHITE), fillc=BLUE, align=center())
ws.row_dimensions[3].height = 34
ws.freeze_panes = "I4"

dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True)
dv.error = "أدخل رقماً موجباً فقط"; dv.errorTitle = "قيمة غير صحيحة"
ws.add_data_validation(dv)

DATA_START = 4
for idx,(ax,name,unit,pol,agg,tgt,ttxt,fmt) in enumerate(KPIS):
    r = DATA_START + idx
    style_cell(ws.cell(r,1,idx+1), f=font(10,True), align=center())
    style_cell(ws.cell(r,2,AX[ax]), f=font(10), align=right(True))
    style_cell(ws.cell(r,3,name), f=font(10), align=right(True))
    style_cell(ws.cell(r,4,unit), align=center())
    style_cell(ws.cell(r,5,pol), align=center())
    style_cell(ws.cell(r,6,agg), align=center())
    gc = ws.cell(r,7, tgt if tgt is not None else "")
    style_cell(gc, fmt=FMT[fmt] if (tgt is not None) else None, align=center())
    style_cell(ws.cell(r,8,ttxt), f=font(9,color="595959"), align=center())
    # خلايا الأشهر — إدخال
    for m in range(12):
        cc = ws.cell(r, 9+m)
        style_cell(cc, fillc=GREEN_IN, fmt=FMT[fmt], align=center(False), lock=False)
        dv.add(cc)
    rng = f"$I{r}:$T{r}"
    # YTD — الصف الفارغ يرجع "" ليظهر "—" بدل صفر يؤثر على اللوحة
    ytd = (f'=IF($F{r}="SUM",IF(COUNT({rng})=0,"",SUM({rng})),'
           f'IF($F{r}="AVG",IFERROR(AVERAGE({rng}),""),'
           f'IF($F{r}="LAST",IFERROR(LOOKUP(2,1/({rng}<>""),{rng}),""),"")))')
    style_cell(ws.cell(r,21,ytd), f=font(10,True), fmt=FMT[fmt], align=center())
    # نسبة التحقيق
    ach = (f'=IF($G{r}="","",IF($U{r}="","",'
           f'IF($G{r}=0,IF($U{r}<=0,1,0),'
           f'IF($E{r}="↓",IFERROR($G{r}/$U{r},0),IFERROR($U{r}/$G{r},0)))))')
    style_cell(ws.cell(r,22,ach), f=font(10,True,NAVY), fmt="0%", align=center())
    # الحالة
    st = (f'=IF($V{r}="","—",IF($V{r}>=1,"✅ محقق",IF($V{r}>=0.85,"🟡 قريب","🔴 تحت الهدف")))')
    style_cell(ws.cell(r,23,st), f=font(10,True), align=center())
    ws.row_dimensions[r].height = 26
DATA_END = DATA_START + len(KPIS) - 1
protect(ws)

ACH_RANGE  = f"'{SH_KPI}'!$V${DATA_START}:$V${DATA_END}"
AXIS_RANGE = f"'{SH_KPI}'!$B${DATA_START}:$B${DATA_END}"
ST_RANGE   = f"'{SH_KPI}'!$W${DATA_START}:$W${DATA_END}"

# ---------------- 3) لوحة الإدارة التنفيذية ----------------
ws = wb.create_sheet("لوحة الإدارة التنفيذية")
# نقلها للموضع الثالث
wb.move_sheet("لوحة الإدارة التنفيذية", -(len(wb.sheetnames)-3))
ws.sheet_view.showGridLines = False
for i,w in enumerate([3,30,16,16,16,16],1): ws.column_dimensions[get_column_letter(i)].width = w
ws.merge_cells("B2:F2")
style_cell(ws["B2"], value="درب · لوحة الإدارة التنفيذية 2026", f=font(20,True,WHITE), fillc=NAVY, align=center(False), border=False)
ws.row_dimensions[2].height = 38
# بطاقات مفتاحية
ws.merge_cells("B4:C4"); style_cell(ws["B4"], value="مؤشر الإنجاز العام (الامتياز)", f=font(11,True,WHITE), fillc=BLUE, align=center())
ws.merge_cells("B5:C6");
style_cell(ws["B5"], value=f'=IFERROR(AVERAGE({ACH_RANGE}),0)', f=font(30,True,NAVY), fmt="0%", align=center())
cards = [("✅ محقق","✅ محقق",GREEN_IN),("🟡 قريب","🟡 قريب",GOLD),("🔴 تحت الهدف","🔴 تحت الهدف","F8CBAD")]
cc=4
for title,key,clr in cards:
    L=get_column_letter(cc)
    style_cell(ws[f"{L}4"], value=title, f=font(11,True), fillc=clr, align=center())
    ws[f"{L}5"].value = f'=COUNTIF({ST_RANGE},"{key}")'
    style_cell(ws[f"{L}5"], f=font(22,True,NAVY), align=center())
    ws.merge_cells(f"{L}5:{L}6")
    cc+=1
ws.row_dimensions[4].height=22; ws.row_dimensions[5].height=30

# جدول المحاور
hr=8
style_cell(ws.cell(hr,2,"الأداء حسب المحور — إدارة الامتياز"), f=font(12,True,WHITE), fillc=BLUE, align=center())
ws.merge_cells(start_row=hr,start_column=2,end_row=hr,end_column=4)
style_cell(ws.cell(hr+1,2,"المحور"), f=font(11,True,WHITE), fillc=STEEL, align=center())
style_cell(ws.cell(hr+1,3,"نسبة التحقيق"), f=font(11,True,WHITE), fillc=STEEL, align=center())
style_cell(ws.cell(hr+1,4,"الحالة"), f=font(11,True,WHITE), fillc=STEEL, align=center())
for i,ax in enumerate(AX):
    r=hr+2+i
    style_cell(ws.cell(r,2,ax), f=font(10), align=right())
    ach=f'=IFERROR(AVERAGEIFS({ACH_RANGE},{AXIS_RANGE},"{ax}"),0)'
    style_cell(ws.cell(r,3,ach), f=font(10,True,NAVY), fmt="0%", align=center())
    st=f'=IF(C{r}>=1,"✅ محقق",IF(C{r}>=0.85,"🟡 قريب","🔴 تحت الهدف"))'
    style_cell(ws.cell(r,4,st), f=font(10,True), align=center())

# جدول الإدارات (تغذية موحّدة)
dr=hr+2+len(AX)+1
style_cell(ws.cell(dr,2,"نظرة الإدارات (تغذي نفس الملف)"), f=font(12,True,WHITE), fillc=BLUE, align=center())
ws.merge_cells(start_row=dr,start_column=2,end_row=dr,end_column=4)
deps=[("الامتياز التجاري والمبيعات", f'=IFERROR(AVERAGE({ACH_RANGE}),0)', True),
      ("التشغيل والمحطات","قيد الإضافة",False),
      ("الاستثمار","قيد الإضافة",False),
      ("العقار","قيد الإضافة",False)]
style_cell(ws.cell(dr+1,2,"الإدارة"), f=font(11,True,WHITE), fillc=STEEL, align=center())
style_cell(ws.cell(dr+1,3,"الإنجاز"), f=font(11,True,WHITE), fillc=STEEL, align=center())
style_cell(ws.cell(dr+1,4,"الحالة"), f=font(11,True,WHITE), fillc=STEEL, align=center())
for i,(dep,val,live) in enumerate(deps):
    r=dr+2+i
    style_cell(ws.cell(r,2,dep), f=font(10), align=right())
    if live:
        style_cell(ws.cell(r,3,val), f=font(10,True,NAVY), fmt="0%", align=center())
        style_cell(ws.cell(r,4,f'=IF(C{r}>=1,"✅ محقق",IF(C{r}>=0.85,"🟡 قريب","🔴 تحت الهدف"))'), f=font(10,True), align=center())
    else:
        style_cell(ws.cell(r,3,val), f=font(9,color="808080"), align=center())
        style_cell(ws.cell(r,4,"—"), align=center())
protect(ws)

# ---------------- 5) مؤشرات كل موظف (مرجع) ----------------
ws = wb.create_sheet("الامتياز · الموظفون")
ws.sheet_view.showGridLines = False
for i,w in enumerate([3,22,34,46,12],1): ws.column_dimensions[get_column_letter(i)].width = w
ws.merge_cells("B2:E2")
style_cell(ws["B2"], value="إدارة الامتياز · مؤشرات كل موظف (مرجع الربط)", f=font(15,True,WHITE), fillc=NAVY, align=center(False), border=False)
ws.row_dimensions[2].height=28
EMP = [
 ("علاء عزت","أخصائي امتياز — عقود واستقطاب", [("●","عدد عقود الامتياز الجديدة الموقّعة","🔴 رئيسي"),("●","عدد محطات الامتياز المشغّلة (تراكمي)","🔴 رئيسي"),("●","نسبة تحويل الفرص إلى عقود موقّعة","🔴 رئيسي"),("●","إجمالي مبيعات وقود محطات الامتياز","🔴 رئيسي"),("●","نسبة استدامة أصحاب الامتياز","🔴 رئيسي")]),
 ("أنس أبو طوق","مراقب ميداني — هوية وبنية تحتية", [("●","نسبة جاهزية محطات الامتياز","🔴 رئيسي"),("●","نسبة إغلاق ملاحظات الجودة الميدانية","🔴 رئيسي"),("●","متوسط درجة تقييم جودة المحطة","🔴 رئيسي"),("●","نسبة المحطات المستوفية لمعايير السلامة","🔴 رئيسي"),("●","نسبة التزام المحطات بالهوية المؤسسية","🔴 رئيسي")]),
 ("حمدي عبد الجليل","أخصائي امتياز — زيارات واستقطاب", [("●","نسبة تنفيذ الزيارات الميدانية المجدولة","🔴 رئيسي"),("○","عدد عقود الامتياز الجديدة الموقّعة","🔴 رئيسي"),("○","متوسط درجة تقييم جودة المحطة","🔴 رئيسي"),("●","نسبة مبيعات المنتجات غير البترولية","🟡 متابعة")]),
 ("رغد الجندي","أخصائية امتياز — تسويق ومالية", [("●","إجمالي إيرادات إدارة الامتياز","🔴 رئيسي"),("●","نسبة التحصيل المالي من شركاء الامتياز","🔴 رئيسي"),("●","متوسط تقييم المحطات على خرائط قوقل","🔴 رئيسي"),("●","نسبة الشركاء الملتزمين بمواعيد السداد","🟠 مهم"),("●","عدد الحملات التسويقية الميدانية","🟠 مهم")]),
 ("إدريس الوائلي","أخصائي امتياز — تقارير وتنسيق مالي", [("●","متوسط أيام التحصيل (DSO)","🟠 مهم"),("●","إجمالي المستحقات المتأخرة (+30 يوم)","🟠 مهم"),("●","نسبة إنجاز التقارير الدورية في موعدها","🟠 مهم"),("○","عدد محطات الامتياز المشغّلة","🔴 رئيسي")]),
 ("أحمد رمضان","أخصائي امتياز — Dynamics ومتابعة نظام", [("○","نسبة اكتمال المستندات النظامية قبل التوقيع","🟠 مهم"),("○","نسبة إنجاز التقارير الدورية في موعدها","🟠 مهم")]),
 ("عبدالسلام الوالي","أخصائي امتياز — عمليات ومخزون", [("●","عدد حالات انقطاع الوقود الموثّقة","🔴 رئيسي"),("●","متوسط وقت معالجة البلاغات التشغيلية","🟠 مهم"),("○","نسبة الشركاء الملتزمين بمواعيد السداد","🟠 مهم")]),
 ("المعتصم المصعبي","أخصائي امتياز — عمليات ومبيعات", [("●","متوسط مبيعات المحطة الواحدة شهرياً","🟠 مهم"),("○","إجمالي مبيعات وقود محطات الامتياز","🔴 رئيسي"),("○","نسبة جاهزية محطات الامتياز","🔴 رئيسي")]),
 ("يحيى العدله","أخصائي امتياز — تسجيل وتوثيق", [("●","نسبة اكتمال المستندات النظامية قبل التوقيع","🟠 مهم"),("○","نسبة حل الشكاوى خلال المدة المحددة","🟠 مهم"),("○","متوسط مدة إغلاق عقد الامتياز","🟠 مهم")]),
 ("إسماعيل سندي","أخصائي امتياز — تواصل عملاء وتأجير", [("●","معدل رضا شركاء الامتياز (CSAT)","🔴 رئيسي"),("●","نسبة حل الشكاوى خلال المدة المحددة","🟠 مهم"),("●","نسبة انخفاض الشكاوى التشغيلية والإدارية","🟠 مهم"),("○","نسبة استدامة أصحاب الامتياز","🔴 رئيسي")]),
]
r=4
for nm,role,kpis in EMP:
    ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=5)
    style_cell(ws.cell(r,2,f"{nm}  —  {role}  ·  {len(kpis)} مؤشر"), f=font(12,True,WHITE), fillc=BLUE, align=right())
    r+=1
    for d in ["","الدور","المؤشر","الأولوية"]:
        pass
    style_cell(ws.cell(r,2,"●/○"), f=font(10,True,WHITE), fillc=STEEL, align=center())
    style_cell(ws.cell(r,3,"المؤشر"), f=font(10,True,WHITE), fillc=STEEL, align=center())
    style_cell(ws.cell(r,4,"الأولوية"), f=font(10,True,WHITE), fillc=STEEL, align=center())
    style_cell(ws.cell(r,5,"", ), fillc=STEEL)
    r+=1
    for mark,kpi,pri in kpis:
        style_cell(ws.cell(r,2,mark), f=font(11,True), align=center())
        style_cell(ws.cell(r,3,kpi), f=font(10), align=right(True))
        style_cell(ws.cell(r,4,pri), f=font(10), align=center())
        style_cell(ws.cell(r,5,""))
        r+=1
    r+=1
protect(ws)

# ---------------- نماذج التقييم الوظيفية ----------------
SCALE = [("ممتاز","90% – 100%","تجاوز المستهدف"),("جيد جداً","80% – 89%","حقق المستهدف تقريباً"),
         ("جيد","70% – 79%","حقق معظم المستهدفات"),("مقبول","60% – 69%","يحتاج تحسين"),("ضعيف","أقل من 60%","يحتاج خطة تطوير")]
def eval_form(sheet, role, code, rows):
    ws = wb.create_sheet(sheet)
    ws.sheet_view.showGridLines = False
    for i,w in enumerate([4,40,10,34,14,14,22],1): ws.column_dimensions[get_column_letter(i)].width = w
    ws.merge_cells("A1:G1")
    style_cell(ws["A1"], value=f"إدارة الامتياز التجاري · نموذج تقييم الأداء — {role}", f=font(14,True,WHITE), fillc=NAVY, align=center(False), border=False)
    ws.row_dimensions[1].height=28
    ws.merge_cells("A2:G2")
    style_cell(ws["A2"], value=f"الكود: {code}   ·   فترة القياس: ربع سنوي   ·   🟩 أدخل «نسبة التحقيق» فقط — الدرجة تُحسب تلقائياً", f=font(10,False,NAVY), fillc=GOLD, align=center(False), border=False)
    # بيانات
    meta=[("اسم الموظف:","القسم: الامتياز التجاري"),("المدير المباشر:","الربع / السنة:"),("تاريخ التقييم:","حالة التقييم:")]
    rr=3
    for a,b in meta:
        style_cell(ws.cell(rr,1,a), f=font(10,True), align=right()); ws.merge_cells(start_row=rr,start_column=1,end_row=rr,end_column=1)
        style_cell(ws.cell(rr,2,""), fillc=GREEN_IN, lock=False)
        ws.merge_cells(start_row=rr,start_column=2,end_row=rr,end_column=3)
        style_cell(ws.cell(rr,4,b), f=font(10,True), align=right())
        style_cell(ws.cell(rr,5,""), fillc=GREEN_IN, lock=False); ws.merge_cells(start_row=rr,start_column=5,end_row=rr,end_column=7)
        rr+=1
    # رأس الجدول
    hdr=["#","مؤشر الأداء","الوزن %","المستهدف","نسبة التحقيق","الدرجة المحققة","ملاحظات"]
    for c,h in enumerate(hdr,1):
        style_cell(ws.cell(rr,c,h), f=font(10,True,WHITE), fillc=BLUE, align=center())
    ws.row_dimensions[rr].height=30
    first=rr+1
    dvp = DataValidation(type="decimal", operator="between", formula1="0", formula2="1.2", allow_blank=True)
    dvp.error="أدخل نسبة بين 0% و120%"; dvp.errorTitle="نسبة التحقيق"
    ws.add_data_validation(dvp)
    rr=first
    for i,(name,w,tgt) in enumerate(rows):
        style_cell(ws.cell(rr,1,i+1), f=font(10,True), align=center())
        style_cell(ws.cell(rr,2,name), f=font(10), align=right(True))
        style_cell(ws.cell(rr,3,w/100.0), fmt="0%", align=center())
        style_cell(ws.cell(rr,4,tgt), f=font(9,color="595959"), align=right(True))
        ac=ws.cell(rr,5); style_cell(ac, fillc=GREEN_IN, fmt="0%", align=center(), lock=False); dvp.add(ac)
        style_cell(ws.cell(rr,6,f"=C{rr}*MIN(E{rr},1)"), f=font(10,True,NAVY), fmt="0%", align=center())
        style_cell(ws.cell(rr,7,""), fillc=GREEN_IN, lock=False)
        ws.row_dimensions[rr].height=28
        rr+=1
    last=rr-1
    # المجموع
    style_cell(ws.cell(rr,1,""))
    style_cell(ws.cell(rr,2,"المجموع الكلي"), f=font(11,True,WHITE), fillc=NAVY, align=center())
    style_cell(ws.cell(rr,3,f"=SUM(C{first}:C{last})"), f=font(11,True,WHITE), fillc=NAVY, fmt="0%", align=center())
    style_cell(ws.cell(rr,4,""), fillc=NAVY)
    style_cell(ws.cell(rr,5,"الدرجة:"), f=font(11,True,WHITE), fillc=NAVY, align=center())
    style_cell(ws.cell(rr,6,f"=SUM(F{first}:F{last})"), f=font(13,True,"FFFF00"), fillc=NAVY, fmt="0%", align=center())
    total_cell=f"F{rr}"
    style_cell(ws.cell(rr,7,""), fillc=NAVY)
    rr+=1
    # التقدير
    style_cell(ws.cell(rr,2,"التقدير"), f=font(11,True,WHITE), fillc=BLUE, align=center())
    grade=(f'=IF({total_cell}>=0.9,"ممتاز",IF({total_cell}>=0.8,"جيد جداً",'
           f'IF({total_cell}>=0.7,"جيد",IF({total_cell}>=0.6,"مقبول","ضعيف"))))')
    style_cell(ws.cell(rr,3,grade), f=font(11,True,NAVY), align=center()); ws.merge_cells(start_row=rr,start_column=3,end_row=rr,end_column=4)
    rr+=2
    # سلّم التقييم
    style_cell(ws.cell(rr,2,"سلّم التقييم"), f=font(11,True,WHITE), fillc=BLUE, align=center()); ws.merge_cells(start_row=rr,start_column=2,end_row=rr,end_column=4)
    rr+=1
    for g,p,m in SCALE:
        style_cell(ws.cell(rr,2,g), f=font(10,True), align=center())
        style_cell(ws.cell(rr,3,p), f=font(10), align=center())
        style_cell(ws.cell(rr,4,m), f=font(10), align=right()); ws.merge_cells(start_row=rr,start_column=4,end_row=rr,end_column=4)
        rr+=1
    protect(ws)

eval_form("الامتياز · مدير الإدارة","مدير إدارة الامتياز التجاري","FM-01",[
 ("عدد محطات الامتياز الجديدة",15,"100%+ من المحطات المستهدفة"),
 ("إجمالي مبيعات محطات الامتياز",15,"100%+ من مستهدف المبيعات الربعي"),
 ("استمرارية تزويد الوقود",10,"0 حالات انقطاع"),
 ("نسبة التحصيل المالي",15,"100%+ من المستحقات الربعية"),
 ("الالتزام بالهوية المؤسسية",15,"70%+ من المحطات ملتزمة"),
 ("انخفاض عدد الشكاوى",10,"خفض 20%+ مقارنة بالربع السابق"),
 ("تقييمات خرائط قوقل",10,"4.7 – 5"),
 ("انتظام وجودة التقارير",10,"100% في موعدها"),
])
eval_form("الامتياز · مسؤول الامتياز","مسؤول الامتياز التجاري","FM-02",[
 ("المتابعة مع عملاء الامتياز",30,"متابعة مستمرة 100%"),
 ("عدم انقطاع المحطات من الوقود",35,"0 حالات انقطاع"),
 ("متابعة طلبات وتوريد الوقود",10,"100% متابعة يومية"),
 ("إجمالي مبيعات المحطات",10,"تحقيق هدف المبيعات"),
 ("التحصيل المالي",10,"100% خلال الربع"),
 ("التقارير الدورية الشاملة",5,"100% في مواعيدها"),
])
eval_form("الامتياز · المراقب الميداني","المراقب الميداني","FM-03",[
 ("عدد الزيارات الميدانية المنفذة",30,"100% من الزيارات المجدولة"),
 ("متابعة تطبيق الهوية المؤسسية",20,"100% دون ملاحظات متكررة"),
 ("تقارير الجودة الميدانية والـChecklist",20,"100% رفع التقارير"),
 ("متابعة العلامات التجارية",10,"100% حسب الجدول"),
 ("تطبيق معايير السلامة الميدانية",10,"100%"),
 ("متابعة الشكاوى وإغلاقها",10,"الاستجابة خلال 24 ساعة"),
])
eval_form("الامتياز · مسؤول التعاقدات","مسؤول التعاقدات","FM-04",[
 ("جودة التفاوض مع العملاء",20,"98% توافق مع المستهدفات"),
 ("اكتمال المستندات النظامية قبل التوقيع",15,"100%"),
 ("سرعة إنجاز إجراءات العقد",30,"خلال يومي عمل 100%"),
 ("أرشفة العقود الإلكترونية",10,"خلال 48 ساعة 100%"),
 ("متابعة تنفيذ بنود العقد",10,"100%"),
 ("مدى رضا العملاء",10,"95% رضا"),
 ("تحديث قاعدة بيانات العملاء",5,"100% دوري"),
])
eval_form("الامتياز · الموظف الإداري","الموظف الإداري","FM-05",[
 ("دقة وسرعة إعداد التقارير الدورية",30,"100% في الوقت المحدد"),
 ("متابعة الشكاوى وحالات العملاء",30,"100% حتى الإغلاق"),
 ("متابعة عمولات تأجير الوحدات",20,"100%"),
 ("تحديث بيانات العملاء والمحطات",10,"100% دقة"),
 ("إعداد المراسلات الرسمية",10,"100% في الوقت المحدد"),
])

# حماية هيكل المصنّف
wb.security = WorkbookProtection(workbookPassword=PWD, lockStructure=True)

import os
os.makedirs("الادارات", exist_ok=True)
out = "درب-مؤشرات-الأداء-2026.xlsx"
wb.save(out)
print("saved:", out, os.path.getsize(out), "bytes")
print("sheets:", wb.sheetnames)
