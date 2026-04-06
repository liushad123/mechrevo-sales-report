import pandas as pd
import json

excel_path = r'C:\Users\刘莎\.easyclaw\workspace\mechrevo-sales-report\data.xlsx'
df = pd.read_excel(excel_path)

# 列索引
COL_PROVINCE = 1   # B 列 - 省份
COL_QTY = 13       # N 列 - 数量
COL_SALES = 31     # AF 列 - 销售金额

# 按省份统计
province_dict = {}
for i in range(len(df)):
    prov = str(df.iloc[i, COL_PROVINCE])
    qty = int(df.iloc[i, COL_QTY])
    sales = float(df.iloc[i, COL_SALES])
    
    if prov not in province_dict:
        province_dict[prov] = {'sales': 0, 'qty': 0}
    
    province_dict[prov]['sales'] += sales
    province_dict[prov]['qty'] += qty

# 转换为列表并排序（按销量）
province_list = [
    {'name': k, 'sales': v['sales'], 'qty': v['qty']} 
    for k, v in province_dict.items()
]
province_list.sort(key=lambda x: x['qty'], reverse=True)

print('=' * 60)
print('区域排行数据（按销量排序）')
print('=' * 60)
print(f'\n{"排名":<4} {"省份":<10} {"销量":>10} {"销售额 (万)":>12} {"占比":>8}')
print('-' * 60)

total_qty = sum(p['qty'] for p in province_list)
total_sales = sum(p['sales'] for p in province_list)

for i, p in enumerate(province_list, 1):
    pct = (p['qty'] / total_qty * 100) if total_qty > 0 else 0
    print(f'{i:<4} {p["name"]:<10} {p["qty"]:>10,} {p["sales"]/10000:>12,.2f} {pct:>7.1f}%')

print('-' * 60)
print(f'总计      {total_qty:>10,} {total_sales/10000:>12,.2f}')

# 保存 JSON
result = {
    'total_qty': total_qty,
    'total_sales': total_sales,
    'provinces': province_list
}

with open('province_data.json', 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False, indent=2)

print(f'\n✅ 数据已保存到 province_data.json')

# 输出 TOP10 用于 HTML 更新
print(f'\n📊 TOP10 省份数据（用于 HTML 更新）:')
for i, p in enumerate(province_list[:10], 1):
    pct = (p['qty'] / total_qty * 100)
    print(f'{i}. {p["name"]}: {p["qty"]:,}台，¥{p["sales"]/10000:.2f}万，占比{pct:.1f}%')
