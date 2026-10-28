import pandas as pd
df = pd.read_pickle('all3.pkl')
jul = df[df.date>=pd.Timestamp('2025-07-01')].copy()
jul['Doctor']=jul.doc.map({'ND':'Dr. Dang','MI':'Dr. Iizuka'})
jul['dow']=jul.date.dt.day_name()
NONCLIN={'Surgery Cataract','Surgery Refractive','Surgery IOLX','MEETING'}
clin = jul[~jul.type.isin(NONCLIN)]

print('=== weekday distribution of clinic dates, by doctor ===')
days_by_doc = clin.groupby(['Doctor','dow']).date.nunique().unstack(fill_value=0)
print(days_by_doc)
print()
print('=== all distinct clinic dates per doctor (sorted) ===')
for doc in ['Dr. Dang','Dr. Iizuka']:
    dates = sorted(clin[clin.Doctor==doc].date.unique())
    print(doc, 'n=',len(dates))
