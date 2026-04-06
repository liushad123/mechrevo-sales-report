import pandas as pd

excel_path = r'C:\Users\刘莎\.easyclaw\workspace\mechrevo-sales-report\data.xlsx'
df = pd.read_excel(excel_path)

print('=' * 60)
print('N Column (Quantity) Re-Check')
print('=' * 60)

# N 列索引 13
n_col = df.iloc[:, 13]
print(f'\nN 列名：{df.columns[13]}')
print(f'N 列数据类型：{n_col.dtype}')
print(f'N 列总和：{n_col.sum():,}')
print(f'N 列记录数：{n_col.count():,}')

# 检查月份列（X 列，索引 23）
month_col = df.iloc[:, 23]
print(f'\n月份列名：{df.columns[23]}')
print(f'月份数据：{month_col.unique()}')

# 按月份分组统计
if '月' in str(df.columns[23]):
    month_summary = df.groupby(df.columns[23]).agg({
        13: 'sum',
        31: 'sum'
    }).reset_index()
    month_summary.columns = ['月份', '数量总和', '销售金额总和']
    print('\n按月份统计:')
    print(month_summary.to_string())

# 检查是否有负数或异常值
print(f'\nN 列统计:')
print(f'  最小值：{n_col.min()}')
print(f'  最大值：{n_col.max()}')
print(f'  平均值：{n_col.mean():.2f}')
print(f'  负数个数：{(n_col < 0).sum()}')

# 如果负数，计算净销量
if (n_col < 0).any():
    positive_sum = n_col[n_col > 0].sum()
    negative_sum = n_col[n_col < 0].sum()
    print(f'\n  正数总和：{positive_sum:,}')
    print(f'  负数总和：{negative_sum:,}')
    print(f'  净销量：{positive_sum + negative_sum:,}')
