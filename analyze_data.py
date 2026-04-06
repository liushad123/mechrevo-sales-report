import pandas as pd
import json

# 读取 Excel 文件
excel_path = r'C:\Users\刘莎\.easyclaw\workspace\mechrevo-sales-report\data.xlsx'

df = pd.read_excel(excel_path)

print('=== 文件基本信息 ===')
print(f'总行数：{len(df)}')
print(f'总列数：{len(df.columns)}')

# AF 列 (索引 31) 是销售金额
af_sum = pd.to_numeric(df.iloc[:, 31], errors='coerce').sum()
print(f'\nAF 列 (销售金额) 总和：{af_sum:,.2f}')
print(f'总销售额 (万元)：{af_sum/10000:.2f}万')

# 按系列统计
series_col = 11  # 系列列
series_summary = df.groupby(df.columns[series_col]).agg({
    df.columns[31]: 'sum',  # 销售金额
    df.columns[24]: 'sum'   # 销售数量
}).reset_index()
series_summary.columns = ['系列', '销售金额', '销售数量']
print('\n=== 按系列统计 ===')
print(series_summary.to_string(index=False))

# 按省份统计
province_col = 1  # 省份列
province_summary = df.groupby(df.columns[province_col]).agg({
    df.columns[31]: 'sum',
    df.columns[24]: 'sum'
}).reset_index()
province_summary.columns = ['省份', '销售金额', '销售数量']
print('\n=== 按省份统计 ===')
print(province_summary.to_string(index=False))

# 按产品型号统计
model_col = 12  # 型号列
model_summary = df.groupby(df.columns[model_col]).agg({
    df.columns[31]: 'sum',
    df.columns[24]: 'sum'
}).reset_index()
model_summary.columns = ['型号', '销售金额', '销售数量']
model_summary = model_summary.sort_values('销售数量', ascending=False)
print('\n=== 产品型号 TOP10 ===')
print(model_summary.head(10).to_string(index=False))

# 按月份统计
month_col = 23  # 月份列
month_summary = df.groupby(df.columns[month_col]).agg({
    df.columns[31]: 'sum',
    df.columns[24]: 'sum'
}).reset_index()
month_summary.columns = ['月份', '销售金额', '销售数量']
print('\n=== 按月份统计 ===')
print(month_summary.to_string(index=False))

# 保存分析结果
result = {
    'total_rows': int(len(df)),
    'total_sales': float(af_sum),
    'total_sales_wan': round(af_sum / 10000, 2),
    'series_summary': series_summary.to_dict('records'),
    'province_summary': province_summary.to_dict('records'),
    'model_top10': model_summary.head(10).to_dict('records'),
    'month_summary': month_summary.to_dict('records')
}

with open('C:/Users/刘莎/.easyclaw/workspace/mechrevo-sales-report/data_analysis.json', 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False, indent=2)

print('\n✅ 分析结果已保存到 data_analysis.json')
print(f'\n📊 总销售额：¥{af_sum:,.2f} (¥{af_sum/10000:.2f}万)')
