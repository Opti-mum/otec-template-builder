import pandas as pd
from classify_refractive import is_refractive

df = pd.read_pickle('all3.pkl')
jul = df[df.date>=pd.Timestamp('2025-07-01')].copy()
jul['Doctor']=jul.doc.map({'ND':'Dr. Dang','MI':'Dr. Iizuka'})
jul['Month']=jul.date.dt.strftime('%Y-%m')
consults = jul[jul.type.isin(['Consult MD','Consult OD'])].copy()
consults['refractive']=consults['reason'].apply(is_refractive)
consults['cat_label']=consults.refractive.map({True:'Refractive',False:'Consult Other'})

print('=== By doctor ===')
print(consults.groupby(['Doctor','cat_label']).size().unstack())
print()
print('=== By doctor x month ===')
piv = consults.groupby(['Doctor','Month','cat_label']).size().unstack(fill_value=0)
print(piv.to_string())
