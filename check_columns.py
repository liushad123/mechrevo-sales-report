import pandas as pd

excel_path = r'C:\Users\刘莎\.easyclaw\workspace\mechrevo-sales-report\data.xlsx'
df = pd.read_excel(excel_path, dtype=str)

print('Column Names (0-32):')
for i in range(len(df.columns)):
    print(f'{i}: {df.columns[i]}')

print('\n\nSample Data (First 3 rows):')
for i in range(3):
    print(f'\nRow {i+1}:')
    for j in range(len(df.columns)):
        val = df.iloc[i, j]
        if pd.notna(val):
            print(f'  [{j}] {df.columns[j]}: {val}')
