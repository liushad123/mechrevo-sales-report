import pandas as pd

excel_path = r'C:\Users\刘莎\.easyclaw\workspace\mechrevo-sales-report\data.xlsx'
df = pd.read_excel(excel_path, dtype=str)

print('=== 列名检查 (A-Z + AA-AF) ===')
print('Excel 列字母 -> 索引 -> 列名')
print('')

# Excel 列字母映射
letters = 'A B C D E F G H I J K L M N O P Q R S T U V W X Y Z AA AB AC AD AE AF'.split()

for i, letter in enumerate(letters):
    if i < len(df.columns):
        print(f'{letter:3} -> {i:2} -> {df.columns[i]}')

# 检查 N 列（索引 13）
print('\n=== N 列（数量）数据检查 ===')
if len(df.columns) > 13:
    n_col_name = df.columns[13]
    print(f'N 列索引：13')
    print(f'N 列名：{n_col_name}')
    print(f'N 列前 10 个值：{df.iloc[:, 13].head(10).tolist()}')
    n_data = pd.to_numeric(df.iloc[:, 13], errors='coerce').fillna(0)
    print(f'N 列总和：{n_data.sum():,.0f}')

# 检查 AF 列（索引 31）
print('\n=== AF 列（销售金额）数据检查 ===')
if len(df.columns) > 31:
    af_col_name = df.columns[31]
    print(f'AF 列索引：31')
    print(f'AF 列名：{af_col_name}')
    print(f'AF 列前 10 个值：{df.iloc[:, 31].head(10).tolist()}')
    af_data = pd.to_numeric(df.iloc[:, 31], errors='coerce').fillna(0)
    print(f'AF 列总和：¥{af_data.sum():,.2f}')
