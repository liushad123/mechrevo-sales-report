import pandas as pd

# 读取 Excel 文件
df = pd.read_excel(r'C:\Users\刘莎\.openclaw\workspace\机械革命 - 销售日报 3 月.xlsx')

print('=== 列名 ===')
print(df.columns.tolist())
print(f'\n总列数：{len(df.columns)}')

print('\n=== AF 列 (第 32 列，索引 31) ===')
if len(df.columns) >= 32:
    af_col_name = df.columns[31]
    print(f'AF 列名：{af_col_name}')
    print(f'AF 列数据预览:')
    print(df.iloc[:, 31].head(20))
    print(f'\nAF 列总和：{df.iloc[:, 31].sum()}')
else:
    print('列数不足 32 列')
    print(f'\n最后 5 列:')
    print(df.columns[-5:])

print('\n=== 所有列名 ===')
for i, col in enumerate(df.columns):
    print(f'{i}: {col}')
