import pandas as pd
df = pd.read_excel('merge_precipitation_gotvand.xlsx')
df_selected = df[['Persian Date', 'Unnamed: 5']].copy()
df_selected.columns = ['Persian Date', 'precipitation']
df_selected['precipitation'] = pd.to_numeric(df_selected['precipitation'], errors='coerce')
df_selected['Persian Date'] = df_selected['Persian Date'].astype(str)
df_selected[['Year', 'Month', 'Day']] = df_selected['Persian Date'].str.split('/', expand=True)
df_selected['Year'] = df_selected['Year'].astype(int)
df_selected['Month'] = df_selected['Month'].astype(int)
df_selected['Day'] = df_selected['Day'].astype(int)
monthly_stats = df_selected.groupby(['Year', 'Month']).agg(Monthly_Sum=('precipitation', 'sum'), Days_Count=('precipitation', 'count')).reset_index()
pivot_table = monthly_stats.pivot(index='Year', columns='Month', values='Monthly_Sum')
pivot_table = pivot_table.reindex(columns=range(1, 13))
pivot_table.index.name = 'Year'
pivot_table.columns.name = None
print(pivot_table)
pivot_table.to_excel('monthly_rainfall_matrix.xlsx')
