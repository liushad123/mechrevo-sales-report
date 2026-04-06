import pandas as pd

excel_path = r'C:\Users\刘莎\.easyclaw\workspace\mechrevo-sales-report\data.xlsx'
df = pd.read_excel(excel_path)

print('=' * 60)
print('Checking all numeric columns for quantity')
print('=' * 60)

# 检查所有列
for i in range(len(df.columns)):
    col_name = str(df.columns[i])
    col_data = df.iloc[:, i]
    
    # 如果是数字类型
    if col_data.dtype in ['int64', 'float64']:
        col_sum = col_data.sum()
        col_mean = col_data.mean()
        print(f'\nCol {i} ({col_name}):')
        print(f'  Type: {col_data.dtype}')
        print(f'  Sum: {col_sum:,.0f}')
        print(f'  Mean: {col_mean:.2f}')
        print(f'  Min: {col_data.min():,.0f}')
        print(f'  Max: {col_data.max():,.0f}')
        print(f'  Sample: {col_data.head(5).tolist()}')
