# -*- coding: utf-8 -*-
import pandas as pd
import json
import sys

# 设置输出编码
sys.stdout.reconfigure(encoding='utf-8')

# 读取 Excel 文件
excel_path = r'C:\Users\刘莎\.easyclaw\workspace\mechrevo-sales-report\data.xlsx'
df = pd.read_excel(excel_path, dtype=str)

print('=' * 60)
print('MechRevo Sales Data Analysis Report')
print('=' * 60)

# 基本信息
total_rows = len(df)
print(f'\n[Basic Info]')
print(f'Total Records: {total_rows:,}')
print(f'Total Columns: {len(df.columns)}')

# AF 列 (索引 31) 是销售金额
af_col = 31
af_data = pd.to_numeric(df.iloc[:, af_col], errors='coerce').fillna(0)
total_sales = af_data.sum()
print(f'\n[Sales Amount - Column AF]')
print(f'Total Sales: RMB {total_sales:,.2f}')
print(f'Total Sales (10K): RMB {total_sales/10000:.2f}')

# AE 列 (索引 30) 是销售数量
ae_col = 30
ae_data = pd.to_numeric(df.iloc[:, ae_col], errors='coerce').fillna(0)
total_qty = ae_data.sum()
print(f'\n[Sales Quantity - Column AE]')
print(f'Total Quantity: {total_qty:,.0f} units')

# 系列分析 (列 11)
print(f'\n[By Series - Column 11]')
series_col = 11
series_data = df.groupby(series_col).agg({
    af_col: lambda x: pd.to_numeric(x, errors='coerce').sum(),
    ae_col: lambda x: pd.to_numeric(x, errors='coerce').sum()
}).reset_index()
series_data.columns = ['Series', 'Sales', 'Qty']
series_data = series_data.sort_values('Sales', ascending=False)
for _, row in series_data.iterrows():
    print(f"  {row['Series']}: RMB {row['Sales']:,.0f} ({row['Qty']:,.0f} units)")

# 型号分析 (列 12)
print(f'\n[By Model TOP10 - Column 12]')
model_col = 12
model_data = df.groupby(model_col).agg({
    af_col: lambda x: pd.to_numeric(x, errors='coerce').sum(),
    ae_col: lambda x: pd.to_numeric(x, errors='coerce').sum()
}).reset_index()
model_data.columns = ['Model', 'Sales', 'Qty']
model_data = model_data.sort_values('Qty', ascending=False).head(10)
for i, (_, row) in enumerate(model_data.iterrows(), 1):
    print(f"  {i}. {row['Model']}: {row['Qty']:,.0f} units (RMB {row['Sales']:,.0f})")

# 省份分析 (列 1)
print(f'\n[By Province - Column 1]')
province_col = 1
province_data = df.groupby(province_col).agg({
    af_col: lambda x: pd.to_numeric(x, errors='coerce').sum(),
    ae_col: lambda x: pd.to_numeric(x, errors='coerce').sum()
}).reset_index()
province_data.columns = ['Province', 'Sales', 'Qty']
province_data = province_data.sort_values('Sales', ascending=False)
for _, row in province_data.iterrows():
    print(f"  {row['Province']}: RMB {row['Sales']:,.0f} ({row['Qty']:,.0f} units)")

# 月份分析 (列 23)
print(f'\n[By Month - Column 23]')
month_col = 23
month_data = df.groupby(month_col).agg({
    af_col: lambda x: pd.to_numeric(x, errors='coerce').sum(),
    ae_col: lambda x: pd.to_numeric(x, errors='coerce').sum()
}).reset_index()
month_data.columns = ['Month', 'Sales', 'Qty']
month_data = month_data.sort_values('Month')
for _, row in month_data.iterrows():
    print(f"  Month {row['Month']}: RMB {row['Sales']:,.0f} ({row['Qty']:,.0f} units)")

# 保存 JSON 结果
result = {
    'overview': {
        'total_rows': int(total_rows),
        'total_sales': float(total_sales),
        'total_sales_wan': round(total_sales / 10000, 2),
        'total_qty': int(total_qty)
    },
    'series': [
        {
            'name': str(row['Series']),
            'sales': float(row['Sales']),
            'qty': int(row['Qty'])
        }
        for _, row in series_data.iterrows()
    ],
    'models': [
        {
            'name': str(row['Model']),
            'sales': float(row['Sales']),
            'qty': int(row['Qty'])
        }
        for _, row in model_data.iterrows()
    ],
    'provinces': [
        {
            'name': str(row['Province']),
            'sales': float(row['Sales']),
            'qty': int(row['Qty'])
        }
        for _, row in province_data.iterrows()
    ],
    'months': [
        {
            'month': str(row['Month']),
            'sales': float(row['Sales']),
            'qty': int(row['Qty'])
        }
        for _, row in month_data.iterrows()
    ]
}

with open('data_analysis.json', 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False, indent=2)

print(f'\n[OK] Results saved to data_analysis.json')
print(f'\n[Summary]')
print(f'   Total Sales: RMB {total_sales:,.2f}')
print(f'   Total Quantity: {total_qty:,.0f} units')
print(f'   Total Records: {total_rows:,}')
