import pandas as pd, re

f = '/root/.claude/uploads/354e382d-9485-594d-8f55-386b4aab8b52/e1aec507-ND_MI_Template_Trend_Analysis_2023-2026.xlsx'
xl = pd.ExcelFile(f)
tg = xl.parse('TG 0724-0625')
tg.columns = ['demo', 'roster', 'date', 'time', 'type', 'prov', 'reason']
tg['date'] = pd.to_datetime(tg['date'])

r = tg['reason'].astype(str)
nd_mask = r.str.contains(r'\bND\b', regex=True)
mi_mask = r.str.contains(r'\bMI\b', regex=True)
own_mask = ~nd_mask & ~mi_mask  # Tarn Gill's own patients - excluded

print('TG sheet rows:', len(tg))
print('  tagged ND (Dr. Dang, counted):', nd_mask.sum())
print('  tagged MI (Dr. Iizuka, counted):', mi_mask.sum())
print('  Tarn Gill\'s own (excluded):', own_mask.sum())

nd_rows = tg[nd_mask].copy()
nd_rows['sheet'] = 'TG 0724-0625'
nd_rows['doc'] = 'ND'
nd_rows['sched'] = "Dr. Dang (on Tarn Gill's schedule)"

mi_rows = tg[mi_mask].copy()
mi_rows['sheet'] = 'TG 0724-0625'
mi_rows['doc'] = 'MI'
mi_rows['sched'] = "Dr. Iizuka (on Tarn Gill's schedule)"

add = pd.concat([nd_rows, mi_rows], ignore_index=True)
add['time'] = add['time'].astype(str)

all3 = pd.read_pickle('all3.pkl')
print('\nBefore:', all3.shape, all3.sched.value_counts().to_dict())

combined = pd.concat([all3, add], ignore_index=True)
combined = combined.drop_duplicates(subset=['demo', 'date', 'time', 'type', 'reason', 'sched'])
print('\nAfter:', combined.shape)
print(combined.sched.value_counts())
print(combined.doc.value_counts())

combined.to_pickle('all3.pkl')
print('\nsaved all3.pkl')

# sanity: monthly counts for ND in the gap window now
combined['Month'] = combined['date'].dt.strftime('%Y-%m')
nd_monthly = combined[combined.doc == 'ND'].groupby('Month').size()
print('\nND monthly counts around the gap:')
print(nd_monthly.loc['2024-09':'2025-07'])
mi_monthly = combined[combined.doc == 'MI'].groupby('Month').size()
print('\nMI monthly counts around the gap:')
print(mi_monthly.loc['2024-09':'2025-07'])
