import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

wb=openpyxl.Workbook(); ws=wb.active; ws.title="تكلفة حرق قطاف"
ws.sheet_view.rightToLeft=True; ws.sheet_view.showGridLines=False
ORANGE="E07C16";ORANGE_L="FCE9D2";INK="26262A";GREY="6D6E70";YELLOW="FFF4CC";GREENL="E4F2E7";HEAD="F3EFEA";WHITE="FFFFFF";REDL="FBE9E8"
thin=Side(style="thin",color="D9D7D3");border=Border(left=thin,right=thin,top=thin,bottom=thin)
def c(r,col,v=None,bold=False,size=11,color=INK,fill=None,align="right",fmt=None,bd=True):
    cc=ws.cell(row=r,column=col,value=v);cc.font=Font(name="Arial",bold=bold,size=size,color=color)
    cc.alignment=Alignment(horizontal=align,vertical="center")
    if fill:cc.fill=PatternFill("solid",fgColor=fill)
    if fmt:cc.number_format=fmt
    if bd:cc.border=border
    return cc
for col,w in {"A":3,"B":34,"C":16,"D":16,"E":16,"F":16}.items(): ws.column_dimensions[col].width=w
SAR='#,##0 "﷼"';NUM='#,##0';PCT='0.0%';PCT0='0%'

ws.merge_cells("B2:F2");c(2,2,"درب — تكلفة الربط مع قطاف (حرق النقاط)",bold=True,size=15,color=ORANGE,align="right",bd=False)
ws.merge_cells("B3:F3");c(3,2,"التكلفة = مبيعات التطبيق × نسبة الدفع بنقاط قطاف × نسبة الخصم · عدّل الأصفر",size=10,color=GREY,align="right",bd=False)

# inputs (single)
c(5,2,"① المدخلات القابلة للتعديل",bold=True,size=12,color=WHITE,fill=ORANGE,bd=False);ws.merge_cells("B5:F5")
c(6,2,"نسبة الدفع بنقاط قطاف من المبيعات",size=11);c(6,3,0.05,fill=YELLOW,align="center",fmt=PCT0);ws.merge_cells("C6:F6")
c(7,2,"نسبة خصم قطاف (مؤكّدة 2%)",size=11);c(7,3,0.02,fill=YELLOW,align="center",fmt=PCT0);ws.merge_cells("C7:F7")
REDMP="C6"; DISC="C7"

# growth table
gt=9
c(gt,2,"② الحساب (3 سنوات)",bold=True,size=12,color=WHITE,fill=ORANGE,bd=False);ws.merge_cells(start_row=gt,start_column=2,end_row=gt,end_column=6)
c(gt+1,2,"البند",bold=True,fill=HEAD,align="center");ws.merge_cells(start_row=gt+1,start_column=2,end_row=gt+1,end_column=3)
c(gt+1,4,"السنة 1",bold=True,fill=HEAD,align="center");c(gt+1,5,"السنة 2",bold=True,fill=HEAD,align="center");c(gt+1,6,"السنة 3",bold=True,fill=HEAD,align="center")
# app sales (editable, from expert forecast)
c(gt+2,2,"مبيعات التطبيق السنوية (﷼)",size=11);ws.merge_cells(start_row=gt+2,start_column=2,end_row=gt+2,end_column=3)
for j,v in enumerate([720000000,1606500000,2700000000]): c(gt+2,4+j,v,fill=YELLOW,align="center",fmt=SAR)
APP=gt+2
def crow(r,label,f,fmt,fill):
    c(r,2,label,size=11,bold=True,fill=fill);ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=3)
    for j,col in enumerate(("D","E","F")): c(r,4+j,f(col),align="center",fmt=fmt,fill=fill,bold=True)
crow(gt+3,"قيمة نقاط قطاف المستبدلة (﷼)", lambda col:f"={col}{APP}*{REDMP}", SAR, HEAD)
crow(gt+4,"التكلفة السنوية (الخسارة) ﷼", lambda col:f"={col}{APP}*{REDMP}*{DISC}", SAR, REDL)
crow(gt+5,"نسبة التكلفة من إجمالي المبيعات", lambda col:f"={col}{gt+4}/{col}{APP}", PCT, GREENL)

c(gt+7,2,"ملاحظة: STC يعوّض درب باقي القيمة (~95–98%). التكلفة تقع على جزء النقاط فقط، مقابل الوصول لقاعدة عملاء قطاف (19 مليون).",size=9,color=GREY,align="right",bd=False)
ws.merge_cells(start_row=gt+7,start_column=2,end_row=gt+7,end_column=6)
c(gt+8,2,"⚠️ ثبّت مع قطاف: نسبة الخصم (2% أم 5%؟) وقيمة النقطة (0.25 أم 0.10 ﷼؟).",size=9,color="B23A34",align="right",bd=False)
ws.merge_cells(start_row=gt+8,start_column=2,end_row=gt+8,end_column=6)

path="/tmp/claude-0/-home-user-khadija/05ed6524-f029-51fd-96d9-884d1c665dcf/scratchpad/qitaf_burn_cost.xlsx"
wb.save(path)
for y,app in enumerate([720000000,1606500000,2700000000],1):
    for d in (0.02,0.05):
        print(f"Y{y} @{int(d*100)}%: cost {app*0.05*d:,.0f}")
print("saved",path)
