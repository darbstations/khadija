import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

wb = openpyxl.Workbook(); ws = wb.active; ws.title = "توقعات المدفوعات"
ws.sheet_view.rightToLeft = True; ws.sheet_view.showGridLines = False

ORANGE="E07C16"; ORANGE_L="FCE9D2"; INK="26262A"; GREY="6D6E70"; YELLOW="FFF4CC"; GREENL="E4F2E7"; HEAD="F3EFEA"; WHITE="FFFFFF"
thin=Side(style="thin",color="D9D7D3"); border=Border(left=thin,right=thin,top=thin,bottom=thin)
def c(r,col,v=None,bold=False,size=11,color=INK,fill=None,align="right",fmt=None,bd=True,wrap=False):
    cc=ws.cell(row=r,column=col,value=v); cc.font=Font(name="Arial",bold=bold,size=size,color=color)
    cc.alignment=Alignment(horizontal=align,vertical="center",wrap_text=wrap)
    if fill: cc.fill=PatternFill("solid",fgColor=fill)
    if fmt: cc.number_format=fmt
    if bd: cc.border=border
    return cc
for col,w in {"A":3,"B":34,"C":18,"D":16,"E":3}.items(): ws.column_dimensions[col].width=w
SAR='#,##0 "﷼"'; NUM='#,##0'; PCT='0%'

ws.merge_cells("B2:D2"); c(2,2,"درب — توقعات المبيعات وعدد العمليات (لبوابة الدفع)",bold=True,size=15,color=ORANGE,align="right",bd=False)
ws.merge_cells("B3:D3"); c(3,2,"أرقام توقعية · عدّل الأصفر فقط",size=10,color=GREY,align="right",bd=False)

# INPUTS
c(5,2,"① المدخلات (توقعاتك)",bold=True,size=12,color=WHITE,fill=ORANGE,bd=False); ws.merge_cells("B5:D5")
rows=[("عدد العملاء النشطين المتوقع (السنة 1)",60000,NUM),
      ("عدد عمليات الشحن للعميل / شهر",2,NUM),
      ("متوسط مبلغ الشحنة الواحدة (﷼)",200,SAR),
      ("حصة مدى من إجمالي القيمة",0.75,PCT),
      ("متوسط قيمة عملية مدى (﷼)",180,SAR),
      ("متوسط قيمة عملية فيزا (﷼)",260,SAR)]
r0=6
for i,(lab,val,fmt) in enumerate(rows):
    r=r0+i
    c(r,2,lab,size=11); c(r,3,val,fill=YELLOW,align="center",fmt=fmt); ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=4)
USERS=f"C{r0}"; TPM=f"C{r0+1}"; AVG=f"C{r0+2}"; MADA=f"C{r0+3}"; TKM=f"C{r0+4}"; TKV=f"C{r0+5}"
c(r0+6,2,"حصة فيزا (تُحسب تلقائياً)",size=10,color=GREY); c(r0+6,3,f"=1-{MADA}",align="center",fmt=PCT,fill=WHITE); ws.merge_cells(start_row=r0+6,start_column=3,end_row=r0+6,end_column=4)

# TOTAL SALES
rt=r0+8
c(rt,2,"② إجمالي المبيعات السنوية المتوقعة",bold=True,size=12,color=WHITE,fill=ORANGE,bd=False); ws.merge_cells(start_row=rt,start_column=2,end_row=rt,end_column=4)
TOTAL=f"C{rt+1}"
c(rt+1,2,"إجمالي الشحن السنوي (﷼)",bold=True,fill=ORANGE_L)
c(rt+1,3,f"={USERS}*{TPM}*12*{AVG}",bold=True,align="center",fmt=SAR,fill=ORANGE_L); ws.merge_cells(start_row=rt+1,start_column=3,end_row=rt+1,end_column=4)

# OUTPUTS the bank asked
ro=rt+3
c(ro,2,"③ الأرقام المطلوبة للبنك",bold=True,size=12,color=WHITE,fill=ORANGE,bd=False); ws.merge_cells(start_row=ro,start_column=2,end_row=ro,end_column=4)
c(ro+1,2,"البند",bold=True,fill=HEAD,align="center"); c(ro+1,3,"مدى",bold=True,fill=HEAD,align="center"); c(ro+1,4,"فيزا",bold=True,fill=HEAD,align="center")
# sales
c(ro+2,2,"المبيعات السنوية (﷼)",bold=True,fill=GREENL)
c(ro+2,3,f"={TOTAL}*{MADA}",align="center",fmt=SAR,fill=GREENL,bold=True)
c(ro+2,4,f"={TOTAL}*(1-{MADA})",align="center",fmt=SAR,fill=GREENL,bold=True)
# transactions
c(ro+3,2,"عدد العمليات السنوية",bold=True,fill=GREENL)
c(ro+3,3,f"=({TOTAL}*{MADA})/{TKM}",align="center",fmt=NUM,fill=GREENL,bold=True)
c(ro+3,4,f"=({TOTAL}*(1-{MADA}))/{TKV}",align="center",fmt=NUM,fill=GREENL,bold=True)
# totals col
c(ro+1,5,"",bd=False)
cE=5
ws.column_dimensions["E"].width=16
c(ro+1,5,"الإجمالي",bold=True,fill=HEAD,align="center")
c(ro+2,5,f"={TOTAL}",align="center",fmt=SAR,fill=GREENL,bold=True)
c(ro+3,5,f"=C{ro+3}+D{ro+3}",align="center",fmt=NUM,fill=GREENL,bold=True)

c(ro+5,2,"ملاحظة: مدى أرخص رسوماً (شحن المحفظة ~1.5 ﷼ ثابت)؛ وجّه الشحن لمدى. الأرقام توقعية تُضبط بحجمك الفعلي.",size=9,color=GREY,align="right",bd=False)
ws.merge_cells(start_row=ro+5,start_column=2,end_row=ro+5,end_column=5)

path="/tmp/claude-0/-home-user-khadija/05ed6524-f029-51fd-96d9-884d1c665dcf/scratchpad/payments_forecast.xlsx"
wb.save(path)
# echo computed example
users=100000; tpm=2; avg=200; mada=0.75; tkm=180; tkv=260
total=users*tpm*12*avg
print("example: total", f"{total:,.0f}", "| mada sales", f"{total*mada:,.0f}", "visa sales", f"{total*(1-mada):,.0f}")
print("mada txns", f"{total*mada/tkm:,.0f}", "visa txns", f"{total*(1-mada)/tkv:,.0f}", "total txns", f"{total*mada/tkm+total*(1-mada)/tkv:,.0f}")
print("saved", path)
