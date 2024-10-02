import pandas as pd
import sqlite3

df = pd.read_excel('data/График выходов подразделение 1.xlsx', sheet_name='Сводка', header = 1)
full_df = df[df['Ф.И.О.'].notna()].drop(3, axis=0)
date_df = full_df.iloc[:,6:].isna()
result = pd.DataFrame(columns=['name', 'start_date', 'end_date'])
for i, row in date_df.iterrows():
    row_res = pd.DataFrame(columns=['name', 'start_date', 'end_date'])
    start_date = []
    end_date = []
    if row[row.index[0]]:
        start_date.append(row.index[0])
    else:
        continue
    for col, v in enumerate(row[1:], start=1):
        pred = row[row.index[col-1]]
        if ((v == False)&(pred==False))|((v == True)&(pred==True)):
            continue
        elif (v==True) & (pred==False):
            start_date.append(row.index[col])
        elif (v==False) & (pred==True):
            end_date.append(row.index[col-1])
        else:
            continue
    if row[row.index[-1]]:
        end_date.append(row.index[-1])
    else:
        continue
    row_res['start_date'] = start_date
    row_res['end_date'] = end_date
    row_res['name'] = full_df.loc[i, 'Ф.И.О.']
    result = pd.concat([result, row_res], axis = 0)

db_connection = sqlite3.connect('BdTrainingCenter.db')
result.to_sql('work_schedule', db_connection, if_exists='append', index=False)
db_connection.close()