import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

wb=openpyxl.Workbook(); ws=wb.active; ws.title="توقعات المدفوعات"
ws.sheet_view.rightToLeft=True; ws.sheet_view.showGridLines=False
ORANGE="E07C16";ORANGE_L="FCE9D2";INK="26262A";GREY="6D6E70";YELLOW="FFF4CC";GREENL="E4F2E7";HEAD="F3EFEA";WHITE="FFFFFF";BLUEL="E7F0FA"
thin=Side(style="thin",color="D9D7D3");border=Border(left=thin,right=thin,top=thin,bottom=thin)
def c(r,col,v=None,bold=False,size=11,color=INK,fill=None,align="right",fmt=None,bd=True):
    cc=ws.cell(row=r,column=col,value=v);cc.font=Font(name="Arial",bold=bold,size=size,color=color)
    cc.alignment=Alignment(horizontal=align,vertical="center")
    if fill:cc.fill=PatternFill("solid",fgColor=fill)
    if fmt:cc.number_format=fmt
    if bd:cc.border=border
    return cc
for col,w in {"A":3,"B":34,"C":16,"D":16,"E":16,"F":16}.items(): ws.column_dimensions[col].width=w
SAR='#,##0 "﷼"';NUM='#,##0';PCT='0%'

ws.merge_cells("B2:F2");c(2,2,"درب — توقعات مبيعات التطبيق على كل المحطات (لبوابة الدفع)",bold=True,size=15,color=ORANGE,align="right",bd=False)
ws.merge_cells("B3:F3");c(3,2,"أرقام توقعية · عدّل الأصفر فقط",size=10,color=GREY,align="right",bd=False)

# ① ثوابت
c(5,2,"① الثوابت",bold=True,size=12,color=WHITE,fill=ORANGE,bd=False);ws.merge_cells("B5:F5")
rows=[("متوسط مبيعات المحطة / شهر (وقود+خدمات) ﷼",200000,SAR),
      ("حصة مدى من القيمة",0.75,PCT),
      ("متوسط قيمة عملية مدى (﷼)",120,SAR),
      ("متوسط قيمة عملية فيزا (﷼)",180,SAR)]
r0=6
for i,(lab,val,fmt) in enumerate(rows):
    r=r0+i;c(r,2,lab,size=11);c(r,3,val,fill=YELLOW,align="center",fmt=fmt);ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=6)
MON=f"C{r0}";MADA=f"C{r0+1}";TKM=f"C{r0+2}";TKV=f"C{r0+3}"

# ② سيناريوهات النمو
gt=r0+5
c(gt,2,"② سيناريوهات النمو (3 سنوات)",bold=True,size=12,color=WHITE,fill=ORANGE,bd=False);ws.merge_cells(start_row=gt,start_column=2,end_row=gt,end_column=6)
c(gt+1,2,"البند",bold=True,fill=HEAD,align="center")
c(gt+1,4,"السنة 1",bold=True,fill=HEAD,align="center");c(gt+1,5,"السنة 2",bold=True,fill=HEAD,align="center");c(gt+1,6,"السنة 3",bold=True,fill=HEAD,align="center")
ws.merge_cells(start_row=gt+1,start_column=2,end_row=gt+1,end_column=3)
# editable rows: stations, penetration
def yrow(r,label,vals,fmt,fill=YELLOW,bold=False,bd=True):
    c(r,2,label,size=11,bold=bold);ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=3)
    for j,v in enumerate(vals): c(r,4+j,v,fill=fill,align="center",fmt=fmt,bold=bold,bd=bd)
yrow(gt+2,"عدد المحطات",[173,200,230],NUM)
yrow(gt+3,"نسبة المبيعات عبر التطبيق",[0.40,0.55,0.70],PCT)
ST=gt+2; PEN=gt+3
# computed
def crow(r,label,f_by_col,fmt,fill):
    c(r,2,label,size=11,bold=True,fill=fill);ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=3)
    for j,col in enumerate(("D","E","F")): c(r,4+j,f_by_col(col),align="center",fmt=fmt,fill=fill,bold=True)
crow(gt+4,"مبيعات المحطات السنوية (﷼)", lambda col:f"={col}{ST}*{MON}*12", SAR, HEAD)
APPr=gt+5
crow(APPr,"مبيعات التطبيق السنوية (﷼)", lambda col:f"={col}{ST}*{MON}*12*{col}{PEN}", SAR, ORANGE_L)
crow(gt+6,"— منها مدى (﷼)", lambda col:f"={col}{APPr}*{MADA}", SAR, GREENL)
crow(gt+7,"— منها فيزا (﷼)", lambda col:f"={col}{APPr}*(1-{MADA})", SAR, GREENL)
crow(gt+8,"عدد عمليات مدى", lambda col:f"=({col}{APPr}*{MADA})/{TKM}", NUM, BLUEL)
crow(gt+9,"عدد عمليات فيزا", lambda col:f"=({col}{APPr}*(1-{MADA}))/{TKV}", NUM, BLUEL)
crow(gt+10,"إجمالي العمليات", lambda col:f"={col}{gt+8}+{col}{gt+9}", NUM, BLUEL)

c(gt+12,2,"الأصفر = مُدخلات تعدّلينها (محطات + نسبة اعتماد لكل سنة). الباقي يُحسب تلقائياً. مدى أرخص رسوماً — وجّه الدفع لمدى.",size=9,color=GREY,align="right",bd=False)
ws.merge_cells(start_row=gt+12,start_column=2,end_row=gt+12,end_column=6)

path="/tmp/claude-0/-home-user-khadija/05ed6524-f029-51fd-96d9-884d1c665dcf/scratchpad/payments_forecast.xlsx"
wb.save(path)
for y,(st,pen) in enumerate([(173,.40),(200,.55),(230,.70)],1):
    tot=st*200000*12;app=tot*pen
    print(f"Y{y}: app {app:,.0f} | mada {app*.75:,.0f} visa {app*.25:,.0f} | txns {app*.75/120+app*.25/180:,.0f}")
print("saved",path)
