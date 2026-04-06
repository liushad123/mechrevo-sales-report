import pandas as pd

excel_path = r'C:\Users\刘莎\.easyclaw\workspace\mechrevo-sales-report\data.xlsx'
df = pd.read_excel(excel_path)

print('=' * 60)
print('Excel Column Mapping (A-N-O...AF)')
print('=' * 60)

# Excel 列字母
letters = 'A B C D E F G H I J K L M N O P Q R S T U V W X Y Z AA AB AC AD AE AF'.split()

print('\nColumn Mapping:')
for i, letter in enumerate(letters):
    if i < len(df.columns):
        col_name = str(df.columns[i])
        col_sum = df.iloc[:, i].sum() if df.iloc[:, i].dtype in ['int64', 'float64'] else 'N/A'
        print(f'{letter:3} (idx {i:2}): {col_name:20} | Sum: {col_sum}')

# 特别检查 N 列
print('\n' + '=' * 60)
print('N Column (Index 13) Detailed Check')
print('=' * 60)
n_col = df.iloc[:, 13]
print(f'Column name: {df.columns[13]}')
print(f'Data type: {n_col.dtype}')
print(f'First 10 values: {n_col.head(10).tolist()}')
print(f'Sum: {n_col.sum():,.0f}')
print(f'Count: {n_col.count():,}')
