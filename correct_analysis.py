import pandas as pd
import json

excel_path = r'C:\Users\刘莎\.easyclaw\workspace\mechrevo-sales-report\data.xlsx'
df = pd.read_excel(excel_path)

# 列索引
COL_MODEL = 12     # M 列 - 型号
COL_SERIES = 11    # L 列 - 系列
COL_PROVINCE = 1   # B 列 - 省份
COL_QTY = 13       # N 列 - 数量
COL_PRICE = 30     # AE 列 - 销售单价
COL_SALES = 31     # AF 列 - 销售金额

# 验证：数量 × 单价 = 金额
print('=' * 60)
print('数据验证：N 列 × AE 列 = AF 列')
print('=' * 60)

valid_count = 0
invalid_count = 0
for i in range(min(10, len(df))):
    qty = df.iloc[i, COL_QTY]
    price = df.iloc[i, COL_PRICE]
    sales = df.iloc[i, COL_SALES]
    calc = qty * price
    if calc == sales:
        valid_count += 1
    else:
        invalid_count += 1
        print(f'行{i+1}: {qty} × {price} = {calc}, AF={sales}')

print(f'\n前 10 行验证：{valid_count}行正确，{invalid_count}行错误')

# 总体汇总
total_qty = int(df.iloc[:, COL_QTY].sum())
total_sales = float(df.iloc[:, COL_SALES].sum())
total_rows = len(df)

print(f'\n【核心数据】')
print(f'总销量（N 列总和）：{total_qty:,} 台')
print(f'总销售额（AF 列总和）：¥{total_sales:,.2f} ({total_sales/10000:.2f}万元)')
print(f'总订单数（总行数）：{total_rows:,} 单')
print(f'平均单价：¥{total_sales/total_qty:,.2f}/台')

# 按型号统计
print(f'\n【按型号统计 TOP10】')
model_data = df.groupby(COL_MODEL).agg({
    COL_QTY: 'sum',
    COL_SALES: 'sum'
}).reset_index()
model_data.columns = ['型号', '数量', '金额']
model_data = model_data.sort_values('数量', ascending=False).head(10)

for i, row in model_data.iterrows():
    print(f'{row["型号"]}: {int(row["数量"]):,}台，¥{row["金额"]/10000:.2f}万')

# 按系列统计
print(f'\n【按系列统计】')
series_data = df.groupby(COL_SERIES).agg({
    COL_QTY: 'sum',
    COL_SALES: 'sum'
}).reset_index()
series_data.columns = ['系列', '数量', '金额']
series_data = series_data.sort_values('数量', ascending=False)

for _, row in series_data.iterrows():
    print(f'{row["系列"]}: {int(row["数量"]):,}台，¥{row["金额"]/10000:.2f}万')

# 按省份统计
print(f'\n【按省份统计 TOP10】')
province_data = df.groupby(COL_PROVINCE).agg({
    COL_QTY: 'sum',
    COL_SALES: 'sum'
}).reset_index()
province_data.columns = ['省份', '数量', '金额']
province_data = province_data.sort_values('数量', ascending=False).head(10)

for _, row in province_data.iterrows():
    print(f'{row["省份"]}: {int(row["数量"]):,}台，¥{row["金额"]/10000:.2f}万')

# 保存完整数据
result = {
    'overview': {
        'total_qty': total_qty,
        'total_sales': total_sales,
        'total_sales_wan': round(total_sales / 10000, 2),
        'total_rows': total_rows,
        'avg_price': round(total_sales / total_qty, 2)
    },
    'models': [
        {
            'name': str(row['型号']),
            'qty': int(row['数量']),
            'sales': float(row['金额'])
        }
        for _, row in model_data.iterrows()
    ],
    'series': [
        {
            'name': str(row['系列']),
            'qty': int(row['数量']),
            'sales': float(row['金额'])
        }
        for _, row in series_data.iterrows()
    ],
    'provinces': [
        {
            'name': str(row['省份']),
            'qty': int(row['数量']),
            'sales': float(row['金额'])
        }
        for _, row in province_data.iterrows()
    ]
}

with open('correct_data.json', 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False, indent=2)

print(f'\n✅ 数据已保存到 correct_data.json')
