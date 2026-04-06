import pandas as pd

excel_path = r'C:\Users\刘莎\.easyclaw\workspace\mechrevo-sales-report\data.xlsx'
df = pd.read_excel(excel_path)

print('可能表示销量的列:')
print('=' * 60)

# 检查 N 列和 O 列
print(f'\nN 列 (索引 13): {df.columns[13]}')
print(f'  总和：{df.iloc[:, 13].sum():,}')
print(f'  前 10 个值：{df.iloc[:, 13].head(10).tolist()}')

print(f'\nO 列 (索引 14): {df.columns[14]}')
print(f'  总和：{df.iloc[:, 14].sum():,}')
print(f'  前 10 个值：{df.iloc[:, 14].head(10).tolist()}')

print(f'\nAE 列 (索引 30): {df.columns[30]}')
print(f'  总和：{df.iloc[:, 30].sum():,}')
print(f'  前 10 个值：{df.iloc[:, 30].head(10).tolist()}')

print(f'\nAF 列 (索引 31): {df.columns[31]}')
print(f'  总和：{df.iloc[:, 31].sum():,}')
print(f'  前 10 个值：{df.iloc[:, 31].head(10).tolist()}')

# 验证：销售金额 = 数量 × 单价？
print('\n验证计算 (前 5 行):')
for i in range(5):
    n = df.iloc[i, 13]
    ae = df.iloc[i, 30]
    af = df.iloc[i, 31]
    calc = n * ae
    print(f'行{i+1}: N×AE = {n}×{ae} = {calc}, AF={af}, 匹配={calc==af}')
