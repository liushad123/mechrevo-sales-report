import pandas as pd

excel_path = r'C:\Users\刘莎\.easyclaw\workspace\mechrevo-sales-report\data.xlsx'
df = pd.read_excel(excel_path)

print('=' * 60)
print('N Column (Quantity) Detailed Check')
print('=' * 60)

# N 列是第 14 列 (索引 13)
n_col = df.iloc[:, 13]

print(f'\nColumn 13 name: {df.columns[13]}')
print(f'Column 13 dtype: {n_col.dtype}')
print(f'\nFirst 20 values:')
for i in range(min(20, len(n_col))):
    print(f'  Row {i+1}: {n_col.iloc[i]}')

print(f'\nBasic stats:')
print(f'  Count: {n_col.count()}')
print(f'  Non-null: {n_col.notna().sum()}')
print(f'  Unique values: {n_col.nunique()}')

# 检查是否有非数字
numeric_n = pd.to_numeric(n_col, errors='coerce')
print(f'\nAfter converting to numeric:')
print(f'  Valid numbers: {numeric_n.notna().sum()}')
print(f'  Invalid (NaN): {numeric_n.isna().sum()}')
print(f'  Sum: {numeric_n.sum():,.0f}')
print(f'  Min: {numeric_n.min():,.0f}')
print(f'  Max: {numeric_n.max():,.0f}')
print(f'  Mean: {numeric_n.mean():,.2f}')

# 检查数据分布
print(f'\nValue distribution:')
value_counts = n_col.value_counts().head(20)
for val, count in value_counts.items():
    print(f'  {val}: {count} rows')
