import pandas as pd
import json

excel_path = r'C:\Users\刘莎\.easyclaw\workspace\mechrevo-sales-report\data.xlsx'
df = pd.read_excel(excel_path, dtype=str)

# 列索引
COL_QTY = 13      # 数量
COL_SALES = 31    # 销售金额

# 转换数值
qty_data = pd.to_numeric(df.iloc[:, COL_QTY], errors='coerce').fillna(0)
sales_data = pd.to_numeric(df.iloc[:, COL_SALES], errors='coerce').fillna(0)

# 总计
total_qty = int(qty_data.sum())
total_sales = float(sales_data.sum())
total_rows = int(len(df))

# 按系列统计
series_dict = {}
for i in range(len(df)):
    series = str(df.iloc[i, 11])
    qty = float(qty_data.iloc[i])
    sales = float(sales_data.iloc[i])
    if series not in series_dict:
        series_dict[series] = {'sales': 0, 'qty': 0}
    series_dict[series]['sales'] += sales
    series_dict[series]['qty'] += qty

series_list = [{'name': k, 'sales': v['sales'], 'qty': int(v['qty'])} for k, v in series_dict.items()]
series_list.sort(key=lambda x: x['sales'], reverse=True)

# 按型号统计
model_dict = {}
for i in range(len(df)):
    model = str(df.iloc[i, 12])
    qty = float(qty_data.iloc[i])
    sales = float(sales_data.iloc[i])
    if model not in model_dict:
        model_dict[model] = {'sales': 0, 'qty': 0}
    model_dict[model]['sales'] += sales
    model_dict[model]['qty'] += qty

model_list = [{'name': k, 'sales': v['sales'], 'qty': int(v['qty'])} for k, v in model_dict.items()]
model_list.sort(key=lambda x: x['qty'], reverse=True)
model_top10 = model_list[:10]

# 按省份统计
province_dict = {}
for i in range(len(df)):
    prov = str(df.iloc[i, 1])
    qty = float(qty_data.iloc[i])
    sales = float(sales_data.iloc[i])
    if prov not in province_dict:
        province_dict[prov] = {'sales': 0, 'qty': 0}
    province_dict[prov]['sales'] += sales
    province_dict[prov]['qty'] += qty

province_list = [{'name': k, 'sales': v['sales'], 'qty': int(v['qty'])} for k, v in province_dict.items()]
province_list.sort(key=lambda x: x['sales'], reverse=True)

# 按月份统计
month_dict = {}
for i in range(len(df)):
    month = str(df.iloc[i, 23]).replace('月', '')
    qty = float(qty_data.iloc[i])
    sales = float(sales_data.iloc[i])
    if month not in month_dict:
        month_dict[month] = {'sales': 0, 'qty': 0}
    month_dict[month]['sales'] += sales
    month_dict[month]['qty'] += qty

month_list = [{'month': k, 'sales': v['sales'], 'qty': int(v['qty'])} for k, v in month_dict.items()]
month_list.sort(key=lambda x: int(x['month']) if x['month'].isdigit() else 0)

# 保存 JSON
result = {
    'overview': {
        'total_sales': total_sales,
        'total_sales_wan': round(total_sales / 10000, 2),
        'total_qty': total_qty,
        'total_rows': total_rows
    },
    'series': series_list,
    'models': model_top10,
    'provinces': province_list,
    'months': month_list
}

with open('data_full.json', 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False, indent=2)

print(f'总销售额：{total_sales:,.2f}')
print(f'总销量：{total_qty:,} 台')
print(f'总记录：{total_rows:,} 条')
print('\n已保存到 data_full.json')
