import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L
from openpyxl.chart import BarChart, Reference

df = pd.read_pickle('upcoming.pkl')
df['time'] = pd.to_datetime(df['time'].astype(str), format='%H:%M:%S').dt.time
df['dow'] = pd.to_datetime(df['date']).dt.day_name()
df = df.sort_values(['clinic_day', 'time']).reset_index(drop=True)

DAYS = sorted(df.clinic_day.unique())
DAY_LABELS = {d: f"{d} ({df[df.clinic_day==d].dow.iloc[0][:3]})" for d in DAYS}
ORDER = ['Consult MD', 'Consult OD', 'Follow Up <6 Mo', 'Follow Up >6 Mo',
         'Post Op POM1 (Month)', 'Post Op POW1 (Week)', 'LPI', 'YAG CAPS', 'OCT ONH']
EXTRA = [t for t in sorted(df.type.unique()) if t not in ORDER]
TYPES = ORDER + EXTRA

F = 'Arial'
H = Font(name=F, bold=True, color='FFFFFF'); HF = PatternFill('solid', fgColor='1D5C8F')
B = Font(name=F, bold=True); N = Font(name=F)
SUB = PatternFill('solid', fgColor='EEF2F5'); GREY = Font(name=F, italic=True, color='5B6B78')

wb = Workbook()

# ---- Raw Data ----
raw = wb.active; raw.title = 'Raw Data'
raw.append(['Demo #', 'Name', 'Age', 'Sex', 'Roster', 'Clinic Day', 'Weekday', 'Appt Time', 'Type', 'Provider', 'Reason'])
for r in df.itertuples():
    raw.append([r.demo, r.name, r.age, r.sex, r.roster if isinstance(r.roster, str) else '',
                r.clinic_day, r.dow, r.time, r.type, r.prov, r.reason if isinstance(r.reason, str) else ''])
n = len(df) + 1
for c in raw[1]: c.font = H; c.fill = HF
for row in raw.iter_rows(min_row=2):
    for c in row: c.font = N
    row[7].number_format = 'h:mm'
for i, w in enumerate([9, 22, 6, 6, 9, 15, 11, 11, 20, 16, 48], 1):
    raw.column_dimensions[L(i)].width = w
raw.freeze_panes = 'A2'; raw.auto_filter.ref = f'A1:K{n}'
R = lambda col: f"'Raw Data'!${col}$2:${col}${n}"

def header(ws, row, vals, start=1):
    for j, v in enumerate(vals, start):
        c = ws.cell(row, j, v); c.font = H; c.fill = HF
        c.alignment = Alignment(horizontal='center' if j > 1 else 'left', wrap_text=True, vertical='center')

# ---- Summary by Type ----
s = wb.create_sheet('Summary by Type', 0)
s['A1'] = "Appointment breakdown, Dr. Dang's upcoming clinic days"; s['A1'].font = Font(name=F, bold=True, size=13)
s['A2'] = 'Every number is a COUNTIFS formula against Raw Data — filter that tab by Clinic Day and Type to check any cell.'; s['A2'].font = GREY
header(s, 4, ['Appointment type'] + [DAY_LABELS[d] for d in DAYS] + ['Total'])
r = 5
for t in TYPES:
    s.cell(r, 1, t).font = N
    for j, d in enumerate(DAYS):
        c = s.cell(r, 2 + j, f'=COUNTIFS({R("I")},$A{r},{R("F")},"{d}")')
        c.font = N; c.number_format = '#,##0;-#,##0;"-"'
    c = s.cell(r, 2 + len(DAYS), f'=SUM(B{r}:{L(1+len(DAYS))}{r})'); c.font = N; c.number_format = '#,##0'
    r += 1
end = r
s.cell(end, 1, 'Total rows (incl. combo visits)').font = B; s.cell(end, 1).fill = SUB
for j in range(len(DAYS) + 1):
    col = L(2 + j); c = s.cell(end, 2 + j, f'=SUM({col}5:{col}{end-1})'); c.font = B; c.fill = SUB; c.number_format = '#,##0'
r = end + 1
s.cell(r, 1, 'Unique patients').font = B
for j, d in enumerate(DAYS):
    c = s.cell(r, 2 + j, f'=SUMPRODUCT(({R("F")}="{d}")/COUNTIFS({R("F")},{R("F")},{R("A")},{R("A")}))')
    c.font = B; c.number_format = '#,##0'
c = s.cell(r, 2 + len(DAYS), f'=SUM(B{r}:{L(1+len(DAYS))}{r})'); c.font = B; c.number_format = '#,##0'
uniq_row = r
r += 2
s.cell(r, 1, 'New consults (Consult MD + Consult OD)').font = N
for j, d in enumerate(DAYS):
    c = s.cell(r, 2 + j, f'=COUNTIFS({R("F")},"{d}",{R("I")},"Consult MD")+COUNTIFS({R("F")},"{d}",{R("I")},"Consult OD")')
    c.font = N; c.number_format = '#,##0'
new_row = r; r += 1
s.cell(r, 1, 'Follow-ups (<6 Mo + >6 Mo)').font = N
for j, d in enumerate(DAYS):
    c = s.cell(r, 2 + j, f'=COUNTIFS({R("F")},"{d}",{R("I")},"Follow Up <6 Mo")+COUNTIFS({R("F")},"{d}",{R("I")},"Follow Up >6 Mo")')
    c.font = N; c.number_format = '#,##0'
fu_row = r; r += 1
s.cell(r, 1, 'Post-ops (POM1 + POW1)').font = N
for j, d in enumerate(DAYS):
    c = s.cell(r, 2 + j, f'=COUNTIFS({R("F")},"{d}",{R("I")},"Post Op POM1 (Month)")+COUNTIFS({R("F")},"{d}",{R("I")},"Post Op POW1 (Week)")')
    c.font = N; c.number_format = '#,##0'
po_row = r; r += 1
s.cell(r, 1, 'Lasers (LPI + YAG CAPS)').font = N
for j, d in enumerate(DAYS):
    c = s.cell(r, 2 + j, f'=COUNTIFS({R("F")},"{d}",{R("I")},"LPI")+COUNTIFS({R("F")},"{d}",{R("I")},"YAG CAPS")')
    c.font = N; c.number_format = '#,##0'
laser_row = r; r += 2
s.cell(r, 1, 'Mix share of booked rows (%)').font = Font(name=F, bold=True, italic=True, color='5B6B78'); r += 1
for lab, src in [('New consults', new_row), ('Follow-ups', fu_row), ('Post-ops', po_row), ('Lasers', laser_row)]:
    s.cell(r, 1, lab).font = N
    for j, d in enumerate(DAYS):
        col = L(2 + j)
        c = s.cell(r, 2 + j, f'=IFERROR({col}{src}/{col}{end},0)'); c.font = N; c.number_format = '0.0%'
    r += 1
s.column_dimensions['A'].width = 34
for j in range(len(DAYS) + 1): s.column_dimensions[L(2 + j)].width = 13
s.row_dimensions[4].height = 30
s.freeze_panes = 'B5'

# small bar chart of new/follow-up/post-op/laser mix
ch = BarChart(); ch.type = 'col'; ch.title = 'Booked-visit mix by clinic day'
ch.height = 9; ch.width = 20; ch.y_axis.title = 'Visits'
cats = Reference(s, min_col=2, min_row=4, max_col=1 + len(DAYS), max_row=4)
data = Reference(s, min_col=1, min_row=new_row, max_col=1 + len(DAYS), max_row=laser_row)
ch.add_data(data, titles_from_data=True, from_rows=True)
ch.set_categories(cats)
ch.legend.position = 'b'
s.add_chart(ch, f'A{laser_row + 10}')

# ---- Capacity tab ----
cap = wb.create_sheet('Capacity Check')
cap['A1'] = 'Slot capacity vs. the ~78–82 slot ceiling (45-min lunch, every 5-min slot used)'
cap['A1'].font = Font(name=F, bold=True, size=13)
cap['A2'] = 'Distinct time slots and double-bookings are counted in Python from Raw Data (Excel has no simple per-slot dedupe formula); unique patients and rows above are formulas.'
cap['A2'].font = GREY
header(cap, 4, ['Clinic day', 'Weekday', 'Unique patients', 'Rows (incl. combo)', 'Distinct time slots used', 'Double-booked slots', 'Max patients in one slot'])
slot_stats = {}
for d in DAYS:
    sub = df[df.clinic_day == d]
    uniq_pt_time = sub.drop_duplicates(['demo', 'time'])
    slot_counts = uniq_pt_time.groupby('time').demo.nunique()
    slot_stats[d] = (len(slot_counts), int((slot_counts > 1).sum()), int(slot_counts.max()))
for i, d in enumerate(DAYS, 5):
    row_idx = i - 4  # matches Summary row position (DAYS in same order)
    cap.cell(i, 1, d).font = N
    cap.cell(i, 2, DAY_LABELS[d].split('(')[1].rstrip(')')).font = N
    cap.cell(i, 3, f"=COPY_PLACEHOLDER").font = N  # replaced below
    cap.cell(i, 4, f'=COUNTIF({R("F")},$A{i})').font = N
    slots, dbl, mx = slot_stats[d]
    cap.cell(i, 5, slots).font = N
    cap.cell(i, 6, dbl).font = N
    cap.cell(i, 7, mx).font = N
for i, d in enumerate(DAYS, 5):
    col_letter = L(2 + DAYS.index(d))
    cap.cell(i, 3).value = f"='Summary by Type'!{col_letter}{uniq_row}"
cap_end = 4 + len(DAYS)
for i, w in enumerate([13, 11, 15, 17, 20, 18, 20], 1): cap.column_dimensions[L(i)].width = w
cap.cell(cap_end + 2, 1, 'Reference: with a 45-min lunch and every 5-min slot filled, a single clinic day has about 78–82 bookable slots.').font = GREY

out = '/mnt/user-data/outputs/Upcoming_Clinics_Breakdown_Oct2026.xlsx'
wb.save(out)
print(out)
print('rows', n - 1)
