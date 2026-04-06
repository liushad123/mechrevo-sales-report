import pandas as pd
import json

excel_path = r'C:\Users\刘莎\.easyclaw\workspace\mechrevo-sales-report\data.xlsx'
df = pd.read_excel(excel_path)

print('列名列表:')
for i, col in enumerate(df.columns):
    print(f'{i}: {col}')

# AF 列 (索引 31) 是销售金额
af_sum = pd.to_numeric(df.iloc[:, 31], errors='coerce').sum()
print(f'\n总销售额：¥{af_sum:,.2f} (¥{af_sum/10000:.2f}万)')

# 保存基础数据
result = {
    'total_sales': float(af_sum),
    'total_sales_wan': round(af_sum / 10000, 2)
}

with open('data_analysis.json', 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False, indent=2)

print('✅ 数据已保存')
