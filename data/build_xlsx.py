# -*- coding: utf-8 -*-
"""يبني ملف إكسل شامل لبيانات وخطط القسم التجاري من data.json"""
import json, io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

SRC = 'data.json'
OUT = 'القسم-التجاري-البيانات-والخطط.xlsx'
D = json.load(io.open(SRC, encoding='utf-8'))
ST = D['stations']
MONTHS = D['platform']['months']
IG = D['platform']['integration']

FONT = 'Arial'
INK = '1F2937'
ACC = 'C2410C'      # orange-ish for headers
HDR_FILL = PatternFill('solid', fgColor='F18A2B')
SUB_FILL = PatternFill('solid', fgColor='FDE8D3')
TOT_FILL = PatternFill('solid', fgColor='FFF3CD')
IN_FILL  = PatternFill('solid', fgColor='FFFF00')   # cells to fill in
THIN = Side(style='thin', color='D0D4D9')
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BLUE = Font(name=FONT, size=10, color='0000FF')     # hardcoded input
BLACK = Font(name=FONT, size=10, color=INK)         # formula
GREEN = Font(name=FONT, size=10, color='008000')    # cross-sheet link

def rtl(ws):
    ws.sheet_view.rightToLeft = True

def title(ws, text, ncols, row=1, note=None):
    ws.cell(row=row, column=1, value=text).font = Font(name=FONT, size=14, bold=True, color=INK)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max(ncols,2))
    ws.cell(row=row, column=1).alignment = Alignment(horizontal='right', vertical='center')
    ws.row_dimensions[row].height = 24
    if note:
        ws.cell(row=row+1, column=1, value=note).font = Font(name=FONT, size=9, italic=True, color='6B7280')
        ws.merge_cells(start_row=row+1, start_column=1, end_row=row+1, end_column=max(ncols,2))
        return row+3
    return row+2

def header(ws, row, headers, widths=None):
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = Font(name=FONT, size=10, bold=True, color='FFFFFF')
        c.fill = HDR_FILL
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = BOX
    ws.row_dimensions[row].height = 30
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(row=row+1, column=1)
    return row+1

def put(ws, r, c, v, font=None, fmt=None, wrap=False, fill=None, align='right'):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = font or BLACK
    cell.border = BOX
    cell.alignment = Alignment(horizontal=align, vertical='top', wrap_text=wrap)
    if fmt: cell.number_format = fmt
    if fill: cell.fill = fill
    return cell

MONEY = '#,##0;(#,##0);-'
NUM   = '#,##0;(#,##0);-'
DEC2  = '#,##0.00;(#,##0.00);-'
PCT   = '0.0%'

wb = Workbook()

# ═══════════════════ 1) دليل الملف ═══════════════════
ws = wb.active; ws.title = 'دليل الملف'; rtl(ws)
r = title(ws, 'القسم التجاري — درب · ملف البيانات والخطط',
          4, note='كل الأرقام والخطط لكل منفذ. الخلايا الزرقاء = مُدخلات تُعدَّل · السوداء = معادلات محسوبة · الصفراء = تحتاج تعبئة.')
r = header(ws, r, ['الورقة', 'المحتوى', 'المصدر'], [26, 60, 30])
guide = [
 ('بطاقات المنافذ', 'بيانات كل محطة: الموقع، النوع، المساحة، التشغيل (مضخات/خزان/عمّال/خدمة ذاتية)', 'سجل الأصول + خرائط جوجل'),
 ('شرائح العملاء', 'لكل محطة: الشريحة، ملفها، ما تحتاجه، والعرض المقدّم لها', 'تحليلات درب + دراسة الموقع'),
 ('المحيط التنافسي', 'بطاقة السيناريو، المنافسون، الميزة، الفجوة، البراندات المقترحة', 'خرائط جوجل — سحب فعلي'),
 ('وحدات التأجير', 'كل وحدة: النوع، العدد، الخدمة، النوع الفرعي، الحالة، الإيجار، المساحة', 'جرد المحطة'),
 ('مبيعات المستأجرين', 'مبيعات كل مستأجر شهرياً + نسبة الإيجار ومبيعات/م² (معادلات)', 'تقارير المستأجرين'),
 ('مبيعات الوقود', 'مزيج المنتجات وحجم ولتر وربح كل منتج + الاتجاه الشهري', 'لوحة تحليلات درب'),
 ('فرص B2B', 'الجوار التجاري كأهداف بيع: القناة، القيمة، الحالة، الأولوية', 'خرائط جوجل + المسح الميداني'),
 ('عقود الأساطيل', 'فئات الأساطيل، آلية التعاقد، الجهات المستهدفة، الحجم والربح', 'إدارة الشراكات'),
 ('المكس التسويقي', '٧Ps لكل محطة + توزيع ميزانية الترويج + الحملات', 'إدارة التسويق'),
 ('النموذج المالي', 'الدخل والمصاريف لكل محطة → الصافي وتكلفة اللتر ونقطة التعادل (معادلات)', 'المالية + التشغيل'),
 ('بطاقة الأداء المشتركة', 'المؤشرات المشتركة لكل منفذ: الربح/الزيارة، حصة التطبيق، التغطية', 'محسوبة من الأوراق أعلاه'),
 ('خطة التنفيذ', 'الركائز، قواعد فضّ التعارض، والتسلسل التنفيذي ٣ مراحل', 'استراتيجية التكامل'),
]
for name, desc, src in guide:
    put(ws, r, 1, name, Font(name=FONT, size=10, bold=True, color=INK), wrap=True)
    put(ws, r, 2, desc, wrap=True)
    put(ws, r, 3, src, Font(name=FONT, size=9, color='6B7280'), wrap=True)
    r += 1
r += 1
put(ws, r, 1, 'ملاحظة', Font(name=FONT, size=10, bold=True, color=INK))
put(ws, r, 2, 'الأرقام المالية (الهوامش والإيجارات والمصاريف) تقديرات نموذجية للاعتماد — تُستبدل بالأرقام المعتمدة من المالية.', wrap=True, fill=TOT_FILL)
ws.cell(row=r, column=2).font = Font(name=FONT, size=10, bold=True, color='9C6500')

# ═══════════════════ 2) بطاقات المنافذ ═══════════════════
ws = wb.create_sheet('بطاقات المنافذ'); rtl(ws)
r = title(ws, 'بطاقات المنافذ', 12)
cols = ['الرمز','اسم المحطة','المدينة','نمط الموقع','الحجم','المساحة','خدمة ذاتية',
        'عدد المضخات','سعة الخزان (لتر)','عدد العمّال','الزيارات/شهر','الوصف']
r = header(ws, r, cols, [10,26,12,22,14,14,11,12,15,11,13,60])
for s in ST:
    m, f = s['meta'], s['finance']
    put(ws, r, 1, m['code'], Font(name=FONT, size=10, bold=True, color=INK), align='center')
    put(ws, r, 2, s['station']['name'], BLUE)
    put(ws, r, 3, m['city'], BLUE); put(ws, r, 4, m['type'], BLUE, wrap=True)
    put(ws, r, 5, m['size'], BLUE); put(ws, r, 6, m['area'], BLUE, align='center')
    put(ws, r, 7, m.get('selfService','لا'), BLUE, align='center')
    put(ws, r, 8, f.get('pumps',0), BLUE, NUM, align='center')
    put(ws, r, 9, f.get('tankCapacity',0), BLUE, NUM)
    put(ws, r,10, f.get('workers',0), BLUE, NUM, align='center')
    put(ws, r,11, f.get('visits',0), BLUE, NUM)
    put(ws, r,12, s['station']['tagline'], BLUE, wrap=True)
    ws.row_dimensions[r].height = 46
    r += 1
ws.cell(row=r+1, column=1, value='التشغيل (المضخات/الخزان/العمّال) تقديرات مبدئية — تُعتمد من سجل الأصول والموارد البشرية.').font = Font(name=FONT, size=9, italic=True, color='9C6500')

# ═══════════════════ 3) شرائح العملاء ═══════════════════
ws = wb.create_sheet('شرائح العملاء'); rtl(ws)
r = title(ws, 'شرائح العملاء — مَن نبيع له في كل منفذ', 6)
r = header(ws, r, ['الرمز','الشريحة','الحصة/الوزن','ملف الشريحة','ما تحتاجه','العرض المقدّم'], [10,24,20,46,40,46])
for s in ST:
    for sg in s['segments']:
        put(ws, r, 1, s['meta']['code'], align='center')
        put(ws, r, 2, f"{sg['ic']} {sg['t']}", Font(name=FONT, size=10, bold=True, color=INK), wrap=True)
        put(ws, r, 3, sg.get('share',''), BLUE, wrap=True)
        put(ws, r, 4, sg.get('profile',''), BLUE, wrap=True)
        put(ws, r, 5, sg.get('need',''), BLUE, wrap=True)
        put(ws, r, 6, sg.get('offer',''), BLUE, wrap=True)
        ws.row_dimensions[r].height = 44
        r += 1

# ═══════════════════ 4) المحيط التنافسي ═══════════════════
ws = wb.create_sheet('المحيط التنافسي'); rtl(ws)
r = title(ws, 'المحيط التنافسي وبطاقة السيناريو', 4)
for s in ST:
    put(ws, r, 1, f"{s['meta']['code']} — {s['station']['name']}", Font(name=FONT, size=11, bold=True, color='FFFFFF'), fill=HDR_FILL)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4); r += 1
    for k, v in [('الطريق', s['scenario']['road']), ('المحطة السابقة', s['scenario']['before']),
                 ('المحطة التالية', s['scenario']['after']), ('كثافة المنافسة', s['scenario']['density']),
                 ('الميزة التنافسية', s['advantage']), ('الفجوة', s['gap'])]:
        put(ws, r, 1, k, Font(name=FONT, size=10, bold=True, color=INK), fill=SUB_FILL)
        put(ws, r, 2, v, BLUE, wrap=True)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        ws.row_dimensions[r].height = 30; r += 1
    put(ws, r, 1, 'العلامات في المحطة', Font(name=FONT, size=10, bold=True, color=INK), fill=SUB_FILL)
    put(ws, r, 2, ' | '.join(f"{a}: {b}" for a, b in s['us']), BLUE, wrap=True)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4); ws.row_dimensions[r].height = 30; r += 1
    put(ws, r, 1, 'المنافسون', Font(name=FONT, size=10, bold=True, color=INK), fill=SUB_FILL)
    put(ws, r, 2, ' | '.join(f"{a}: {b}" for a, b in s['them']), BLUE, wrap=True)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4); ws.row_dimensions[r].height = 30; r += 1
    put(ws, r, 1, 'براندات مقترحة', Font(name=FONT, size=10, bold=True, color=INK), fill=SUB_FILL)
    put(ws, r, 2, ' | '.join(f"{a} ({b})" for a, b in s['suggested']), BLUE, wrap=True)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4); ws.row_dimensions[r].height = 30; r += 2
for i, w in enumerate([22, 40, 30, 30], start=1):
    ws.column_dimensions[get_column_letter(i)].width = w

# ═══════════════════ 5) وحدات التأجير ═══════════════════
ws = wb.create_sheet('وحدات التأجير'); rtl(ws)
r = title(ws, 'وحدات التأجير — مكوّنات المحطة', 9)
hdr_row = header(ws, r, ['الرمز','اسم الوحدة/المستأجر','نوع الوحدة','العدد','قسم الخدمة','النوع الفرعي','الحالة','الإيجار/شهر','المساحة م²'],
                 [10,30,16,8,18,18,12,14,12])
r = hdr_row
lease_rows = {}
for s in ST:
    for t in s['tenants']['items']:
        put(ws, r, 1, s['meta']['code'], align='center')
        put(ws, r, 2, t['name'], BLUE, wrap=True)
        put(ws, r, 3, t['unitType'], BLUE, align='center')
        put(ws, r, 4, t.get('count',1), BLUE, NUM, align='center')
        put(ws, r, 5, t['serviceCat'], BLUE, align='center')
        put(ws, r, 6, t['subType'], BLUE, align='center')
        put(ws, r, 7, t['status'], BLUE, align='center')
        put(ws, r, 8, t.get('rent',0), BLUE, MONEY)
        put(ws, r, 9, t.get('area',0), BLUE, NUM, align='center')
        lease_rows.setdefault(s['meta']['code'], []).append(r)
        r += 1
last = r-1
put(ws, r, 1, 'الإجمالي', Font(name=FONT, size=10, bold=True, color=INK), fill=TOT_FILL)
put(ws, r, 4, f'=SUM(D{hdr_row}:D{last})', Font(name=FONT, size=10, bold=True, color=INK), NUM, fill=TOT_FILL, align='center')
put(ws, r, 8, f'=SUM(H{hdr_row}:H{last})', Font(name=FONT, size=10, bold=True, color=INK), MONEY, fill=TOT_FILL)
put(ws, r, 9, f'=SUM(I{hdr_row}:I{last})', Font(name=FONT, size=10, bold=True, color=INK), NUM, fill=TOT_FILL, align='center')
for c in (2,3,5,6,7): put(ws, r, c, '', fill=TOT_FILL)
LEASE_HDR, LEASE_LAST = hdr_row, last

# ═══════════════════ 6) مبيعات المستأجرين (شهرية) ═══════════════════
ws = wb.create_sheet('مبيعات المستأجرين'); rtl(ws)
r = title(ws, 'مبيعات المستأجرين الشهرية', 4+len(MONTHS),
          note='نسبة الإيجار = الإيجار ÷ مبيعات آخر شهر (الصحّي ≤ ١٥٪) · مبيعات/م² = مبيعات آخر شهر ÷ المساحة')
cols = ['الرمز','المستأجر','قسم الخدمة'] + MONTHS + ['الإيجار/شهر','المساحة م²','نسبة الإيجار','مبيعات/م²']
r = header(ws, r, cols, [10,26,16] + [13]*len(MONTHS) + [13,11,13,12])
start = r
nm = len(MONTHS)
lastM = get_column_letter(3+nm)          # آخر شهر
colRent = get_column_letter(4+nm); colArea = get_column_letter(5+nm)
colRatio = get_column_letter(6+nm); colSpm = get_column_letter(7+nm)
for s in ST:
    for t in s['tenants']['items']:
        if t['status'] != 'مؤجّرة':
            continue
        put(ws, r, 1, s['meta']['code'], align='center')
        put(ws, r, 2, t['name'], BLUE, wrap=True)
        put(ws, r, 3, t['serviceCat'], BLUE, align='center')
        sm = t.get('salesM') or [0]*nm
        for i in range(nm):
            put(ws, r, 4+i, sm[i] if i < len(sm) else 0, BLUE, MONEY)
        put(ws, r, 4+nm, t.get('rent',0), BLUE, MONEY)
        put(ws, r, 5+nm, t.get('area',0), BLUE, NUM, align='center')
        put(ws, r, 6+nm, f'=IFERROR({colRent}{r}/{lastM}{r},"")', BLACK, PCT)
        put(ws, r, 7+nm, f'=IFERROR({lastM}{r}/{colArea}{r},"")', BLACK, NUM)
        r += 1
endr = r-1
put(ws, r, 1, 'الإجمالي', Font(name=FONT, size=10, bold=True, color=INK), fill=TOT_FILL)
put(ws, r, 2, '', fill=TOT_FILL); put(ws, r, 3, '', fill=TOT_FILL)
for i in range(nm):
    cl = get_column_letter(4+i)
    put(ws, r, 4+i, f'=SUM({cl}{start}:{cl}{endr})', Font(name=FONT, size=10, bold=True, color=INK), MONEY, fill=TOT_FILL)
put(ws, r, 4+nm, f'=SUM({colRent}{start}:{colRent}{endr})', Font(name=FONT, size=10, bold=True, color=INK), MONEY, fill=TOT_FILL)
put(ws, r, 5+nm, f'=SUM({colArea}{start}:{colArea}{endr})', Font(name=FONT, size=10, bold=True, color=INK), NUM, fill=TOT_FILL, align='center')
put(ws, r, 6+nm, f'=IFERROR({colRent}{r}/{lastM}{r},"")', Font(name=FONT, size=10, bold=True, color=INK), PCT, fill=TOT_FILL)
put(ws, r, 7+nm, '', fill=TOT_FILL)
TEN_TOTAL_ROW, TEN_LASTM_COL = r, lastM

# ═══════════════════ 7) مبيعات الوقود ═══════════════════
ws = wb.create_sheet('مبيعات الوقود'); rtl(ws)
r = title(ws, 'مبيعات الوقود — مزيج المنتجات', 6,
          note='لتر المنتج = الحصة × إجمالي اللترات · ربح المنتج = لتر المنتج × هامش الوقود')
r = header(ws, r, ['الرمز','المنتج','الحصة %','إجمالي اللترات/شهر','هامش ر.س/لتر','لتر المنتج/شهر','ربح المنتج/شهر'],
           [10,16,11,18,14,17,17])
for s in ST:
    f = s['finance']
    for mx in s['fuel']['mix']:
        put(ws, r, 1, s['meta']['code'], align='center')
        put(ws, r, 2, mx[0], BLUE, align='center')
        put(ws, r, 3, (mx[1] or 0)/100.0, BLUE, PCT)
        put(ws, r, 4, f.get('liters',0), BLUE, NUM)
        put(ws, r, 5, f.get('margin',0), BLUE, DEC2)
        put(ws, r, 6, f'=C{r}*D{r}', BLACK, NUM)
        put(ws, r, 7, f'=F{r}*E{r}', BLACK, MONEY)
        r += 1
r += 1
put(ws, r, 1, 'الاتجاه الشهري (مليون لتر)', Font(name=FONT, size=11, bold=True, color=INK)); r += 1
r = header(ws, r, ['الرمز'] + MONTHS, [10] + [13]*nm)
for s in ST:
    put(ws, r, 1, s['meta']['code'], align='center')
    for i, v in enumerate(s['fuel']['monthly'][:nm]):
        put(ws, r, 2+i, v, BLUE, DEC2)
    r += 1

# ═══════════════════ 8) فرص B2B ═══════════════════
ws = wb.create_sheet('فرص B2B'); rtl(ws)
r = title(ws, 'فرص المبيعات — الجوار التجاري (B2B)', 6)
hdr = header(ws, r, ['الرمز','الجهة','النشاط / الوصف','قناة البيع','القيمة ر.س/شهر','الحالة','الأولوية'],
             [10,30,48,18,15,13,11])
r = hdr
for s in ST:
    for o in s['pipeline']['items']:
        put(ws, r, 1, s['meta']['code'], align='center')
        put(ws, r, 2, o['name'], BLUE, wrap=True)
        put(ws, r, 3, o['type'], BLUE, wrap=True)
        put(ws, r, 4, o['channel'], BLUE, align='center')
        put(ws, r, 5, o['value'], BLUE, MONEY)
        put(ws, r, 6, o['status'], BLUE, align='center')
        put(ws, r, 7, o['pri'], BLUE, align='center')
        ws.row_dimensions[r].height = 30
        r += 1
put(ws, r, 1, 'إجمالي خط الأنابيب', Font(name=FONT, size=10, bold=True, color=INK), fill=TOT_FILL)
for c in (2,3,4,6,7): put(ws, r, c, '', fill=TOT_FILL)
put(ws, r, 5, f'=SUM(E{hdr}:E{r-1})', Font(name=FONT, size=10, bold=True, color=INK), MONEY, fill=TOT_FILL)
put(ws, r+2, 1, 'المتعاقد', Font(name=FONT, size=10, bold=True, color=INK))
put(ws, r+2, 2, f'=SUMIF(F{hdr}:F{r-1},"متعاقد",E{hdr}:E{r-1})', BLACK, MONEY)
put(ws, r+3, 1, 'فرص مفتوحة', Font(name=FONT, size=10, bold=True, color=INK))
put(ws, r+3, 2, f'=E{r}-B{r+2}', BLACK, MONEY)

# ═══════════════════ 9) عقود الأساطيل ═══════════════════
ws = wb.create_sheet('عقود الأساطيل'); rtl(ws)
r = title(ws, 'عقود وقود الأساطيل (B2B)', 7)
hdr = header(ws, r, ['الرمز','فئة الأسطول','المنتج','آلية التعاقد','الجهات المستهدفة','الحجم لتر/شهر','هامش ر.س/لتر','ربح متوقّع/شهر','الحالة'],
             [10,26,14,60,50,15,13,16,12])
r = hdr
for s in ST:
    mg = s['finance'].get('margin', 0)
    for f in s['fleet']['items']:
        put(ws, r, 1, s['meta']['code'], align='center')
        put(ws, r, 2, f['seg'], Font(name=FONT, size=10, bold=True, color=INK), wrap=True)
        put(ws, r, 3, f['product'], BLUE, align='center')
        put(ws, r, 4, f['mechanism'], BLUE, wrap=True)
        put(ws, r, 5, f['targets'], BLUE, wrap=True)
        put(ws, r, 6, f['liters'], BLUE, NUM)
        put(ws, r, 7, mg, BLUE, DEC2)
        put(ws, r, 8, f'=F{r}*G{r}', BLACK, MONEY)
        put(ws, r, 9, f['status'], BLUE, align='center')
        ws.row_dimensions[r].height = 58
        r += 1
put(ws, r, 1, 'الإجمالي', Font(name=FONT, size=10, bold=True, color=INK), fill=TOT_FILL)
for c in (2,3,4,5,7,9): put(ws, r, c, '', fill=TOT_FILL)
put(ws, r, 6, f'=SUM(F{hdr}:F{r-1})', Font(name=FONT, size=10, bold=True, color=INK), NUM, fill=TOT_FILL)
put(ws, r, 8, f'=SUM(H{hdr}:H{r-1})', Font(name=FONT, size=10, bold=True, color=INK), MONEY, fill=TOT_FILL)

# ═══════════════════ 10) المكس التسويقي ═══════════════════
ws = wb.create_sheet('المكس التسويقي'); rtl(ws)
r = title(ws, 'المكس التسويقي (٧Ps) — حسب نمط الموقع', 3)
r = header(ws, r, ['الرمز','العنصر','القرار الاستراتيجي'], [10,20,100])
for s in ST:
    for m in s['marketing']['mix']:
        put(ws, r, 1, s['meta']['code'], align='center')
        put(ws, r, 2, f"{m['ic']} {m['p']}", Font(name=FONT, size=10, bold=True, color=INK), align='center')
        put(ws, r, 3, m['d'], BLUE, wrap=True)
        ws.row_dimensions[r].height = 42
        r += 1
r += 1
put(ws, r, 1, 'توزيع ميزانية الترويج', Font(name=FONT, size=11, bold=True, color=INK)); r += 1
bh = header(ws, r, ['الرمز','البند','الحصة %'], [10,34,12]); r = bh
budget_start = {}
for s in ST:
    budget_start[s['meta']['code']] = r
    for b in s['marketing']['budget']:
        put(ws, r, 1, s['meta']['code'], align='center')
        put(ws, r, 2, b[0], BLUE)
        put(ws, r, 3, (b[1] or 0)/100.0, BLUE, PCT)
        r += 1
    put(ws, r, 1, '', fill=TOT_FILL)
    put(ws, r, 2, 'المجموع (يجب أن يساوي ١٠٠٪)', Font(name=FONT, size=10, bold=True, color=INK), fill=TOT_FILL)
    put(ws, r, 3, f'=SUM(C{budget_start[s["meta"]["code"]]}:C{r-1})', Font(name=FONT, size=10, bold=True, color=INK), PCT, fill=TOT_FILL)
    r += 2
put(ws, r, 1, 'الحملات', Font(name=FONT, size=11, bold=True, color=INK)); r += 1
ch = header(ws, r, ['الرمز','الحملة','النوع','المموّل','التكلفة ر.س','الأثر %','الحالة'], [10,36,22,14,14,11,12]); r = ch
for s in ST:
    for c in s['marketing']['campaigns']:
        put(ws, r, 1, s['meta']['code'], align='center')
        put(ws, r, 2, c['name'], BLUE, wrap=True)
        put(ws, r, 3, c['type'], BLUE, wrap=True)
        put(ws, r, 4, c['funder'], BLUE, align='center')
        put(ws, r, 5, c['cost'], BLUE, MONEY)
        put(ws, r, 6, (c.get('lift') or 0)/100.0, BLUE, PCT)
        put(ws, r, 7, c['status'], BLUE, align='center')
        r += 1
cend = r-1
put(ws, r, 2, 'إجمالي تكلفة الحملات', Font(name=FONT, size=10, bold=True, color=INK), fill=TOT_FILL)
put(ws, r, 5, f'=SUM(E{ch}:E{cend})', Font(name=FONT, size=10, bold=True, color=INK), MONEY, fill=TOT_FILL)
put(ws, r+1, 2, 'منها على ميزانية التسويق', Font(name=FONT, size=10, bold=True, color=INK))
put(ws, r+1, 5, f'=SUMIF(D{ch}:D{cend},"التسويق",E{ch}:E{cend})', BLACK, MONEY)
put(ws, r+2, 2, 'مموّل من الشركاء/المستأجرين', Font(name=FONT, size=10, bold=True, color=INK))
put(ws, r+2, 5, f'=E{r}-E{r+1}', BLACK, MONEY)

# ═══════════════════ 11) النموذج المالي ═══════════════════
ws = wb.create_sheet('النموذج المالي'); rtl(ws)
r = title(ws, 'النموذج المالي لكل منفذ (شهرياً)', 2+len(ST),
          note='المُدخلات بالأزرق · النتائج بالأسود (معادلات). عدّل أي مُدخل ليعاد الحساب.')
hdr_r = r
put(ws, r, 1, 'البند', Font(name=FONT, size=10, bold=True, color='FFFFFF'), fill=HDR_FILL, align='center')
put(ws, r, 2, 'الوحدة', Font(name=FONT, size=10, bold=True, color='FFFFFF'), fill=HDR_FILL, align='center')
for j, s in enumerate(ST):
    put(ws, r, 3+j, f"{s['meta']['code']}\n{s['station']['name']}", Font(name=FONT, size=10, bold=True, color='FFFFFF'), fill=HDR_FILL, align='center')
ws.row_dimensions[r].height = 34
put(ws, r, 3+len(ST), 'المحفظة', Font(name=FONT, size=10, bold=True, color='FFFFFF'), fill=HDR_FILL, align='center')
ws.column_dimensions['A'].width = 30; ws.column_dimensions['B'].width = 12
for j in range(len(ST)+1):
    ws.column_dimensions[get_column_letter(3+j)].width = 17
ws.freeze_panes = ws.cell(row=r+1, column=3)
r += 1
def section(label):
    global r
    put(ws, r, 1, label, Font(name=FONT, size=10, bold=True, color=INK), fill=SUB_FILL)
    for j in range(len(ST)+2): put(ws, r, 2+j, '', fill=SUB_FILL)
    r += 1
def inrow(label, unit, key, fmt=MONEY):
    global r
    put(ws, r, 1, label); put(ws, r, 2, unit, Font(name=FONT, size=9, color='6B7280'), align='center')
    for j, s in enumerate(ST):
        put(ws, r, 3+j, s['finance'].get(key, 0), BLUE, fmt)
    r += 1
    return r-1
def frow(label, unit, tmpl, fmt=MONEY, bold=False, fill=None, total=True):
    global r
    fnt = Font(name=FONT, size=10, bold=True, color=INK) if bold else BLACK
    put(ws, r, 1, label, fnt, fill=fill); put(ws, r, 2, unit, Font(name=FONT, size=9, color='6B7280'), align='center', fill=fill)
    for j in range(len(ST)):
        col = get_column_letter(3+j)
        put(ws, r, 3+j, tmpl.format(c=col), fnt, fmt, fill=fill)
    if total:
        cols = [get_column_letter(3+j) for j in range(len(ST))]
        put(ws, r, 3+len(ST), '=' + '+'.join(f'{c}{r}' for c in cols), fnt, fmt, fill=fill)
    else:
        put(ws, r, 3+len(ST), '', fill=fill)
    r += 1
    return r-1

section('⛽ الوقود')
R_LIT = inrow('اللترات', 'لتر/شهر', 'liters', NUM)
R_MRG = inrow('هامش الوقود', 'ر.س/لتر', 'margin', DEC2)
R_FUELINC = frow('دخل الوقود', 'ر.س', '={c}'+str(R_LIT)+'*{c}'+str(R_MRG), bold=True)
section('🏪 التأجير والخدمات')
R_UNITS = inrow('عدد الوحدات', 'وحدة', 'units', NUM)
R_OCC = inrow('نسبة الإشغال', '%', 'occ', NUM)
R_RENT = inrow('متوسط الإيجار للوحدة', 'ر.س', 'rent')
R_STORE = inrow('دخل المتجر/الخدمات', 'ر.س', 'store')
R_RENTINC = frow('دخل التأجير', 'ر.س', '={c}'+str(R_UNITS)+'*{c}'+str(R_OCC)+'/100*{c}'+str(R_RENT), bold=True)
R_NONFUEL = frow('إجمالي غير الوقود', 'ر.س', '={c}'+str(R_RENTINC)+'+{c}'+str(R_STORE), bold=True)
section('💸 المصاريف')
R_SAL = inrow('رواتب', 'ر.س', 'salaries')
R_OPR = inrow('إيجار/تشغيل', 'ر.س', 'oprent')
R_UTL = inrow('كهرباء وماء', 'ر.س', 'utils')
R_MNT = inrow('صيانة وأخرى', 'ر.س', 'maint')
R_EXP = frow('إجمالي المصاريف', 'ر.س', '=SUM({c}'+str(R_SAL)+':{c}'+str(R_MNT)+')', bold=True)
section('🏭 التشغيل')
R_WRK = inrow('عدد العمّال', 'عامل', 'workers', NUM)
R_TANK = inrow('سعة الخزان', 'لتر', 'tankCapacity', NUM)
R_PUMP = inrow('عدد المضخات', 'مضخة', 'pumps', NUM)
R_VIS = inrow('الزيارات', 'زيارة/شهر', 'visits', NUM)
section('📊 النتائج')
R_INC = frow('إجمالي الدخل', 'ر.س', '={c}'+str(R_FUELINC)+'+{c}'+str(R_NONFUEL), bold=True)
R_NET = frow('صافي الربح/الخسارة', 'ر.س', '={c}'+str(R_INC)+'-{c}'+str(R_EXP), bold=True, fill=TOT_FILL)
R_PPV = frow('الربح لكل زيارة', 'ر.س', '=IFERROR({c}'+str(R_NET)+'/{c}'+str(R_VIS)+',"")', DEC2, bold=True, total=False)
frow('تكلفة التشغيل/لتر', 'ر.س', '=IFERROR({c}'+str(R_EXP)+'/{c}'+str(R_LIT)+',"")', '#,##0.000', total=False)
frow('تكلفة الرواتب/عامل', 'ر.س', '=IFERROR({c}'+str(R_SAL)+'/{c}'+str(R_WRK)+',"")', MONEY, total=False)
frow('تغطية غير الوقود للمصاريف', '%', '=IFERROR({c}'+str(R_NONFUEL)+'/{c}'+str(R_EXP)+',"")', PCT, total=False)
frow('أيام تغطية الخزان', 'يوم', '=IFERROR({c}'+str(R_TANK)+'/({c}'+str(R_LIT)+'/30),"")', DEC2, total=False)
frow('لتر لكل مضخة', 'لتر/شهر', '=IFERROR({c}'+str(R_LIT)+'/{c}'+str(R_PUMP)+',"")', NUM, total=False)
frow('نقطة التعادل (لترات)', 'لتر', '=IFERROR(MAX(0,{c}'+str(R_EXP)+'-{c}'+str(R_NONFUEL)+')/{c}'+str(R_MRG)+',"")', NUM, total=False)
# الربح/الزيارة للمحفظة
cols = [get_column_letter(3+j) for j in range(len(ST))]
tot_col = get_column_letter(3+len(ST))
put(ws, R_PPV, 3+len(ST), f'=IFERROR({tot_col}{R_NET}/({"+".join(f"{c}{R_VIS}" for c in cols)}),"")',
    Font(name=FONT, size=10, bold=True, color=INK), DEC2)
FIN_SHEET, FIN_ROWS = 'النموذج المالي', dict(net=R_NET, vis=R_VIS, exp=R_EXP, nonfuel=R_NONFUEL, ppv=R_PPV)

# ═══════════════════ 12) بطاقة الأداء المشتركة ═══════════════════
ws = wb.create_sheet('بطاقة الأداء المشتركة'); rtl(ws)
r = title(ws, 'بطاقة الأداء المشتركة — المؤشرات التي تُقاس عليها الإدارات معاً', 7,
          note='القيم مرتبطة بورقة «النموذج المالي» — أي تعديل هناك ينعكس هنا تلقائياً.')
r = header(ws, r, ['الرمز','المنفذ','الزيارات/شهر','الربح/الزيارة','حصة التطبيق','مستهدف التطبيق','تغطية غير الوقود','ربح المنفذ'],
           [10,26,14,14,13,14,16,16])
sc_start = r
app_target = (IG.get('appTarget') or 5)/100.0
for j, s in enumerate(ST):
    col = get_column_letter(3+j)
    app = next((p[1] for p in s['payment'] if 'تطبيق' in p[0]), 0)
    put(ws, r, 1, s['meta']['code'], align='center')
    put(ws, r, 2, s['station']['name'], BLUE)
    put(ws, r, 3, f"='{FIN_SHEET}'!{col}{FIN_ROWS['vis']}", GREEN, NUM)
    put(ws, r, 4, f"='{FIN_SHEET}'!{col}{FIN_ROWS['ppv']}", GREEN, DEC2)
    put(ws, r, 5, (app or 0)/100.0, BLUE, PCT)
    put(ws, r, 6, app_target, BLUE, PCT)
    put(ws, r, 7, f"=IFERROR('{FIN_SHEET}'!{col}{FIN_ROWS['nonfuel']}/'{FIN_SHEET}'!{col}{FIN_ROWS['exp']},\"\")", GREEN, PCT)
    put(ws, r, 8, f"='{FIN_SHEET}'!{col}{FIN_ROWS['net']}", GREEN, MONEY)
    r += 1
sc_end = r-1
put(ws, r, 1, '', fill=TOT_FILL)
put(ws, r, 2, 'المحفظة', Font(name=FONT, size=10, bold=True, color=INK), fill=TOT_FILL)
put(ws, r, 3, f'=SUM(C{sc_start}:C{sc_end})', Font(name=FONT, size=10, bold=True, color=INK), NUM, fill=TOT_FILL)
put(ws, r, 4, f'=IFERROR(H{r}/C{r},"")', Font(name=FONT, size=10, bold=True, color=INK), DEC2, fill=TOT_FILL)
put(ws, r, 5, '', fill=TOT_FILL); put(ws, r, 6, '', fill=TOT_FILL); put(ws, r, 7, '', fill=TOT_FILL)
put(ws, r, 8, f'=SUM(H{sc_start}:H{sc_end})', Font(name=FONT, size=10, bold=True, color=INK), MONEY, fill=TOT_FILL)
PORT_ROW = r
r += 2
put(ws, r, 1, 'أثر رفع غير الوقود ١ ر.س لكل زيارة', Font(name=FONT, size=10, bold=True, color=INK))
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
put(ws, r, 3, f'=C{PORT_ROW}', BLACK, MONEY, fill=TOT_FILL)
put(ws, r, 4, 'ر.س/شهر على مستوى المحفظة', Font(name=FONT, size=9, italic=True, color='6B7280'))
ws.cell(row=r, column=3).comment = Comment('كل زيارة إضافية بريال واحد من غير الوقود = إجمالي الزيارات الشهرية ريالاً.', 'القسم التجاري')

# ═══════════════════ 13) خطة التنفيذ ═══════════════════
ws = wb.create_sheet('خطة التنفيذ'); rtl(ws)
r = title(ws, 'خطة التنفيذ — التكامل بين الإدارات', 4)
put(ws, r, 1, IG.get('principle',''), Font(name=FONT, size=12, bold=True, color='9C6500'), fill=TOT_FILL, align='center')
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4); ws.row_dimensions[r].height = 28; r += 1
put(ws, r, 1, IG.get('unit',''), Font(name=FONT, size=10, color=INK), wrap=True)
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4); ws.row_dimensions[r].height = 30; r += 2
put(ws, r, 1, 'الركائز الهيكلية', Font(name=FONT, size=11, bold=True, color=INK)); r += 1
r = header(ws, r, ['الركيزة','التفصيل'], [30,90])
for p in IG['pillars']:
    put(ws, r, 1, p['t'], Font(name=FONT, size=10, bold=True, color=INK), wrap=True)
    put(ws, r, 2, p['d'], BLUE, wrap=True); ws.row_dimensions[r].height = 32; r += 1
r += 1
put(ws, r, 1, 'دورة النمو', Font(name=FONT, size=11, bold=True, color=INK)); r += 1
r = header(ws, r, ['#','المرحلة','الأثر'], [6,30,60])
for i, c in enumerate(IG['cycle'], start=1):
    put(ws, r, 1, i, align='center'); put(ws, r, 2, f"{c['ic']} {c['t']}", Font(name=FONT, size=10, bold=True, color=INK))
    put(ws, r, 3, c['d'], BLUE, wrap=True); r += 1
r += 1
put(ws, r, 1, 'قواعد فضّ التعارض', Font(name=FONT, size=11, bold=True, color=INK)); r += 1
r = header(ws, r, ['التعارض','القاعدة الحاكمة'], [46,74])
for a, b in IG['rules']:
    put(ws, r, 1, a, Font(name=FONT, size=10, color='B91C1C'), wrap=True)
    put(ws, r, 2, b, Font(name=FONT, size=10, color='15803D'), wrap=True)
    ws.row_dimensions[r].height = 32; r += 1
r += 1
put(ws, r, 1, 'التسلسل التنفيذي', Font(name=FONT, size=11, bold=True, color=INK)); r += 1
r = header(ws, r, ['المرحلة','التركيز','المبادرات','المبرّر','المسؤول','الحالة'], [16,20,50,34,20,14])
for p in IG['phases']:
    put(ws, r, 1, p['p'], Font(name=FONT, size=10, bold=True, color=ACC), align='center')
    put(ws, r, 2, p['t'], Font(name=FONT, size=10, bold=True, color=INK))
    put(ws, r, 3, p['d'], BLUE, wrap=True)
    put(ws, r, 4, p['why'], Font(name=FONT, size=9, color='6B7280'), wrap=True)
    put(ws, r, 5, '', BLUE, fill=IN_FILL)
    put(ws, r, 6, '', BLUE, fill=IN_FILL, align='center')
    ws.row_dimensions[r].height = 40; r += 1
put(ws, r+1, 1, 'الخلايا الصفراء (المسؤول/الحالة) تُعبَّأ عند الاعتماد.', Font(name=FONT, size=9, italic=True, color='9C6500'))

for s in wb.worksheets:
    s.sheet_properties.tabColor = 'F18A2B'
wb.save(OUT)
print('saved:', OUT)
print('sheets:', [s.title for s in wb.worksheets])
