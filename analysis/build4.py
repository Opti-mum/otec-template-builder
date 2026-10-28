import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L
from openpyxl.chart import LineChart, Reference
from classify_refractive import is_refractive

df = pd.read_pickle('all3.pkl')
jul = df[df.date >= pd.Timestamp('2025-07-01')].copy()
jul['time'] = pd.to_datetime(jul.time.astype(str), format='%H:%M:%S').dt.time
jul['dow'] = jul.date.dt.day_name()
jul = jul.sort_values(['date', 'time']).reset_index(drop=True)
jul['Doctor'] = jul.doc.map({'ND': 'Dr. Dang', 'MI': 'Dr. Iizuka'})
jul['Month'] = jul.date.dt.strftime('%Y-%m')

CORR = {'Macumira': 'Follow Up <6 Mo'}
NONCLIN = {'Surgery Cataract', 'Surgery Refractive', 'Surgery IOLX', 'MEETING'}
jul['type_c'] = jul.type.replace(CORR)
jul['refractive'] = False
is_consult = jul.type_c.isin(['Consult MD', 'Consult OD'])
jul.loc[is_consult, 'refractive'] = jul.loc[is_consult, 'reason'].apply(is_refractive)
jul['row_label'] = jul.type_c
jul.loc[is_consult & jul.refractive, 'row_label'] = 'Consult - Refractive'
jul.loc[is_consult & ~jul.refractive, 'row_label'] = 'Consult - Other'

ORDER = ['Consult - Refractive', 'Consult - Other', 'Follow Up <6 Mo', 'Follow Up >6 Mo',
         'Post Op POM1 (Month)', 'Post Op POW1 (Week)', 'Post Op SD/POD1 (Day)',
         'LPI', 'LPI OD / OS', 'YAG CAPS', 'YAG CAPS OD / OS', 'SLT', 'LDD',
         'Scleral Lens Consult/Dispense', 'Scleral Lens Diagnostic', 'Virtual Call',
         'Surgery Counselling', 'Surgery IOLX', 'Surgery Cataract', 'Surgery Refractive']
EXTRA = [t for t in sorted(jul.row_label.unique()) if t not in ORDER]
TYPES = ORDER + EXTRA
months = sorted(jul.Month.unique())
docs = ['Dr. Dang', 'Dr. Iizuka']

F = 'Arial'
H = Font(name=F, bold=True, color='FFFFFF'); HF = PatternFill('solid', fgColor='1D5C8F')
B = Font(name=F, bold=True); N = Font(name=F)
SUB = PatternFill('solid', fgColor='EEF2F5'); GREY = Font(name=F, italic=True, color='5B6B78')
AMB = PatternFill('solid', fgColor='FBF0D9'); WARN = Font(name=F, italic=True, color='9A6200')

wb = Workbook()

# ---- Raw Data ----
raw = wb.active; raw.title = 'Raw Data'
raw.append(['Demo #', 'Doctor', 'Appt Date', 'Appt Time', 'Weekday', 'Month', 'Type (row label)',
            'Reason', 'Roster', 'Type (as booked in Oscar)', 'Clinic visit?'])
for r in jul.itertuples():
    raw.append([r.demo, r.Doctor, r.date.date(), r.time, r.dow, r.Month, r.row_label,
                r.reason if isinstance(r.reason, str) else '',
                r.roster if isinstance(r.roster, str) else '',
                r.type, 'No' if r.type in NONCLIN else 'Yes'])
n = len(jul) + 1
for c in raw[1]: c.font = H; c.fill = HF
for row in raw.iter_rows(min_row=2):
    for c in row: c.font = N
    row[2].number_format = 'yyyy-mm-dd'; row[3].number_format = 'h:mm'
for i, w in enumerate([9, 12, 12, 9, 11, 9, 22, 55, 13, 24, 12], 1):
    raw.column_dimensions[L(i)].width = w
raw.freeze_panes = 'A2'; raw.auto_filter.ref = f'A1:K{n}'
R = lambda col: f"'Raw Data'!${col}$2:${col}${n}"

def header(ws, row, vals, start=1):
    for j, v in enumerate(vals, start):
        c = ws.cell(row, j, v); c.font = H; c.fill = HF
        c.alignment = Alignment(horizontal='center' if j > 1 else 'left', wrap_text=True, vertical='center')

def typerows(ws, start, ncols, fn):
    r = start
    for t in TYPES:
        ws.cell(r, 1, t).font = N
        for j in range(ncols):
            c = ws.cell(r, 2 + j, fn(r, j)); c.font = N; c.number_format = '#,##0;-#,##0;"-"'
        r += 1
    return r

def totals(ws, end, first, ncols, label='Total'):
    ws.cell(end, 1, label).font = B; ws.cell(end, 1).fill = SUB
    for j in range(ncols):
        col = L(2 + j); c = ws.cell(end, 2 + j, f'=SUM({col}{first}:{col}{end-1})'); c.font = B; c.fill = SUB; c.number_format = '#,##0'

# ---- Summary by Type (both doctors, all months combined) ----
s = wb.create_sheet('Summary by Type', 0)
s['A1'] = "Appointment breakdown by type, July 2025 – Sept 2026 (new consults split into Refractive vs. Consult Other)"
s['A1'].font = Font(name=F, bold=True, size=13)
s['A2'] = 'Refractive = CAT / Cataract / RLE is the earliest-mentioned clinical reason in the Reason text (a primary reason, not a secondary item in a list like "Glc/cat/retina"). COUNTIFS against Raw Data — every cell is checkable.'
s['A2'].font = GREY
s['A3'] = 'This window starts July 2025, when both doctors moved fully onto their own OSCAR schedules (no optometrist-schedule merges needed).'
s['A3'].font = WARN
s.merge_cells('B5:B5'); s.merge_cells('C5:C5')
header(s, 5, ['Appointment type', 'Dr. Dang', 'Dr. Iizuka', 'Total'])
def fs(r, j):
    if j in (0, 1):
        return f'=COUNTIFS({R("G")},$A{r},{R("B")},"{docs[j]}")'
    return f'=B{r}+C{r}'
end = typerows(s, 6, 3, fs); totals(s, end, 6, 3, 'Total rows')
s.cell(end + 1, 1, 'Rows in Raw Data (should match)').font = GREY
c = s.cell(end + 1, 4, f'=COUNTA({R("A")})'); c.font = GREY; c.number_format = '#,##0'
s.column_dimensions['A'].width = 26
for j in range(3): s.column_dimensions[L(2 + j)].width = 13
s.freeze_panes = 'B6'

# ---- monthly per doctor ----
for dname, title in [('Dr. Dang', 'Dr. Dang by Month'), ('Dr. Iizuka', 'Dr. Iizuka by Month')]:
    ws = wb.create_sheet(title)
    ws['A1'] = f'{dname}: appointments by type and month, July 2025 – Sept 2026'; ws['A1'].font = Font(name=F, bold=True, size=13)
    header(ws, 3, ['Appointment type'] + months + ['Total'])
    nm = len(months)
    def fm(r, j, dname=dname):
        if j < nm:
            return f'=COUNTIFS({R("G")},$A{r},{R("B")},"{dname}",{R("F")},{L(2+j)}$3)'
        return f'=SUM(B{r}:{L(1+nm)}{r})'
    end = typerows(ws, 4, nm + 1, fm); totals(ws, end, 4, nm + 1)
    ws.column_dimensions['A'].width = 24
    for j in range(nm + 1): ws.column_dimensions[L(2 + j)].width = 9
    ws.freeze_panes = 'B4'

# ---- ND+MI Combined by Month ----
comb = wb.create_sheet('ND+MI Combined by Month')
comb['A1'] = 'Dr. Dang + Dr. Iizuka combined: appointments by type and month, July 2025 – Sept 2026'
comb['A1'].font = Font(name=F, bold=True, size=13)
comb['A2'] = 'Each cell sums both doctors for that type and month. Rows with no Dr. Iizuka volume (e.g. SLT, LDD) still show Dr. Dang\'s count.'
comb['A2'].font = GREY
header(comb, 4, ['Appointment type'] + months + ['Total'])
nm = len(months)
def fc(r, j):
    if j < nm:
        return f'=COUNTIFS({R("G")},$A{r},{R("F")},{L(2+j)}$4)'
    return f'=SUM(B{r}:{L(1+nm)}{r})'
end_c = typerows(comb, 5, nm + 1, fc); totals(comb, end_c, 5, nm + 1)
comb.column_dimensions['A'].width = 24
for j in range(nm + 1): comb.column_dimensions[L(2 + j)].width = 9
comb.freeze_panes = 'B5'

# ---- Clinic Days & Volume ----
clin = jul[~jul.type.isin(NONCLIN)]
days = clin.groupby(['Doctor', 'Month']).date.nunique()
rows_by_month = jul.groupby(['Doctor', 'Month']).size()

cd = wb.create_sheet('Clinic Days & Volume')
cd['A1'] = 'Clinic days worked vs. appointment volume, by month'
cd['A1'].font = Font(name=F, bold=True, size=13)
cd['A2'] = 'Clinic days = distinct dates with at least one clinic visit (Raw Data column K = "Yes"). Weeks with zero clinic days for a doctor are flagged below as likely away-weeks, inferred from gaps in the data — please confirm against actual vacation/away dates.'
cd['A2'].font = GREY
header(cd, 4, ['Month'] + [f'{d} clinic days' for d in docs] + [f'{d} visits' for d in docs]
       + [f'{d} visits/day' for d in docs] + [f'{d} away weeks (inferred)' for d in docs])
for i, m in enumerate(months, 5):
    cd.cell(i, 1, m).font = N
    for k, d in enumerate(docs):
        cdays = int(days.get((d, m), 0))
        cd.cell(i, 2 + k, cdays).font = N
        col_v = L(4 + k)
        cd.cell(i, 4 + k, f"=COUNTIFS({R('F')},$A{i},{R('B')},\"{d}\")").font = N
        c = cd.cell(i, 6 + k, f'=IFERROR({col_v}{i}/{L(2+k)}{i},0)'); c.font = N; c.number_format = '0.0'
last = 4 + len(months)
for j, col in enumerate('BCDEFG'):
    c = cd.cell(last + 1, 2 + j, f'=SUM({col}5:{col}{last})'); c.font = B; c.fill = SUB; c.number_format = '#,##0'
cd.cell(last + 1, 1, 'Total').font = B; cd.cell(last + 1, 1).fill = SUB
for i, w in enumerate([11, 15, 16, 13, 14, 15, 15, 20, 20], 1): cd.column_dimensions[L(i)].width = w
cd.row_dimensions[4].height = 36
cd.freeze_panes = 'B5'

# inferred away-weeks (zero clinic days in a week) -> list per doctor, and per-month counts
def away_weeks(doc_dates):
    s_ = pd.Series(pd.to_datetime(list(doc_dates)))
    wk = s_.dt.to_period('W-SUN')
    present = set(wk)
    full_range = pd.period_range(jul.date.min(), jul.date.max(), freq='W-SUN')
    return [w for w in full_range if w not in present]

nd_away = away_weeks(clin[clin.Doctor == 'Dr. Dang'].date.unique())
mi_away = away_weeks(clin[clin.Doctor == 'Dr. Iizuka'].date.unique())

def away_count_for_month(away_list, month_str):
    return sum(1 for w in away_list if w.start_time.strftime('%Y-%m') == month_str
               or w.end_time.strftime('%Y-%m') == month_str)

for i, m in enumerate(months, 5):
    nd_c = away_count_for_month(nd_away, m)
    mi_c = away_count_for_month(mi_away, m)
    c1 = cd.cell(i, 8, nd_c if nd_c else ''); c1.font = N
    c2 = cd.cell(i, 9, mi_c if mi_c else ''); c2.font = N
    if nd_c:
        cd.cell(i, 2).fill = AMB; cd.cell(i, 4).fill = AMB; cd.cell(i, 6).fill = AMB; cd.cell(i, 8).fill = AMB
    if mi_c:
        cd.cell(i, 3).fill = AMB; cd.cell(i, 5).fill = AMB; cd.cell(i, 7).fill = AMB; cd.cell(i, 9).fill = AMB

r0 = last + 3
cd.cell(r0, 1, 'Shaded months contain at least one inferred away-week (zero clinic days that week) — please confirm against actual vacation dates.').font = WARN; r0 += 2
cd.cell(r0, 1, 'Weeks with zero Dr. Dang clinic days (inferred away — please confirm)').font = WARN; r0 += 1
for w in nd_away:
    cd.cell(r0, 1, str(w)).font = N; r0 += 1
r0 += 1
cd.cell(r0, 1, 'Weeks with zero Dr. Iizuka clinic days — Dr. Iizuka\'s clinics at OTEC are occasional (not every Tue/Fri), so most of these are expected, not vacation').font = WARN; r0 += 1
for w in mi_away:
    cd.cell(r0, 1, str(w)).font = N; r0 += 1

# chart: visits/day over time, both doctors
cats = Reference(cd, min_col=1, min_row=5, max_row=last)
ch = LineChart(); ch.title = 'Appointment visits per clinic day, by month'; ch.height = 9; ch.width = 26
ch.y_axis.title = 'Visits per clinic day'; ch.x_axis.title = 'Month'
for col, colr in [(6, '1D5C8F'), (7, 'C26A0A')]:
    ch.add_data(Reference(cd, min_col=col, min_row=4, max_row=last), titles_from_data=True)
    s_ = ch.series[-1]; s_.graphicalProperties.line.solidFill = colr; s_.graphicalProperties.line.width = 22000
ch.set_categories(cats); ch.x_axis.tickLblSkip = 2; ch.legend.position = 'b'
ch.x_axis.delete = False; ch.y_axis.delete = False
ch.y_axis.majorGridlines = __import__('openpyxl').chart.axis.ChartLines()
cd.add_chart(ch, 'I4')

# second chart: raw clinic days per month (to see fewer-days effect directly)
ch2 = LineChart(); ch2.title = 'Clinic days worked, by month'; ch2.height = 9; ch2.width = 26
ch2.y_axis.title = 'Clinic days'; ch2.x_axis.title = 'Month'
for col, colr in [(2, '1D5C8F'), (3, 'C26A0A')]:
    ch2.add_data(Reference(cd, min_col=col, min_row=4, max_row=last), titles_from_data=True)
    s_ = ch2.series[-1]; s_.graphicalProperties.line.solidFill = colr; s_.graphicalProperties.line.width = 22000
ch2.set_categories(cats); ch2.x_axis.tickLblSkip = 2; ch2.legend.position = 'b'
ch2.x_axis.delete = False; ch2.y_axis.delete = False
ch2.y_axis.majorGridlines = __import__('openpyxl').chart.axis.ChartLines()
cd.add_chart(ch2, 'I24')

out = '/mnt/user-data/outputs/ND_MI_Refractive_Breakdown_Jul2025_onward.xlsx'
wb.save(out)
print(out, 'rows', n - 1, 'months', len(months))
