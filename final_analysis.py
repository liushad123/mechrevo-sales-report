import pandas as pd
import json

excel_path = r'C:\Users\刘莎\.easyclaw\workspace\mechrevo-sales-report\data.xlsx'
df = pd.read_excel(excel_path, dtype=str)

# 正确的列索引
COL_QTY = 13       # N 列 - 数量
COL_SALES = 31     # AF 列 - 销售金额
COL_SERIES = 11    # L 列 - 系列
COL_MODEL = 12     # M 列 - 型号
COL_PROVINCE = 1   # B 列 - 省份
COL_MONTH = 23     # X 列 - 月份

# 转换数值
qty_data = pd.to_numeric(df.iloc[:, COL_QTY], errors='coerce').fillna(0)
sales_data = pd.to_numeric(df.iloc[:, COL_SALES], errors='coerce').fillna(0)

# 总计
total_qty = int(qty_data.sum())
total_sales = float(sales_data.sum())
total_rows = int(len(df))

print('=' * 60)
print('机械革命销售数据完整分析（从原始 Excel 抓取）')
print('=' * 60)
print(f'\n【核心数据】')
print(f'Total Sales: RMB {total_sales:,.2f} ({total_sales/10000:.2f} wan) - From AF column')
print(f'Total Qty: {total_qty:,} units - From N column')
print(f'Total Orders: {total_rows:,} - Total rows')

# 按系列统计
series_dict = {}
for i in range(len(df)):
    series = str(df.iloc[i, COL_SERIES])
    qty = float(qty_data.iloc[i])
    sales = float(sales_data.iloc[i])
    if series not in series_dict:
        series_dict[series] = {'sales': 0, 'qty': 0}
    series_dict[series]['sales'] += sales
    series_dict[series]['qty'] += qty

series_list = [{'name': k, 'sales': v['sales'], 'qty': int(v['qty'])} for k, v in series_dict.items()]
series_list.sort(key=lambda x: x['qty'], reverse=True)

print(f'\n[Series TOP5]')
for i, s in enumerate(series_list[:5], 1):
    print(f'{i}. {s["name"]}: {s["qty"]:,} units, RMB {s["sales"]:,.0f}')

# 按省份统计
province_dict = {}
for i in range(len(df)):
    prov = str(df.iloc[i, COL_PROVINCE])
    qty = float(qty_data.iloc[i])
    sales = float(sales_data.iloc[i])
    if prov not in province_dict:
        province_dict[prov] = {'sales': 0, 'qty': 0}
    province_dict[prov]['sales'] += sales
    province_dict[prov]['qty'] += qty

province_list = [{'name': k, 'sales': v['sales'], 'qty': int(v['qty'])} for k, v in province_dict.items()]
province_list.sort(key=lambda x: x['qty'], reverse=True)

print(f'\n[Province TOP5]')
for i, p in enumerate(province_list[:5], 1):
    print(f'{i}. {p["name"]}: {p["qty"]:,} units, RMB {p["sales"]:,.0f}')

# 保存完整数据
result = {
    'overview': {
        'total_sales': total_sales,
        'total_sales_wan': round(total_sales / 10000, 2),
        'total_qty': total_qty,
        'total_rows': total_rows,
        'data_source': 'N 列 (数量) + AF 列 (销售金额)'
    },
    'series': series_list,
    'provinces': province_list
}

with open('final_data.json', 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False, indent=2)

print(f'\n[OK] Saved to final_data.json')
print(f'\n[Core Data for HTML]')
print(f'   Total Sales: RMB {total_sales/10000:.2f} wan')
print(f'   Total Qty: {total_qty:,} units')
print(f'   Total Orders: {total_rows:,}')
