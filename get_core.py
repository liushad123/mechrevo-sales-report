import pandas as pd
import json

excel_path = r'C:\Users\刘莎\.easyclaw\workspace\mechrevo-sales-report\data.xlsx'
df = pd.read_excel(excel_path, dtype=str)

# AF 列 (31) = 销售金额，AE 列 (30) = 销售数量
af_data = pd.to_numeric(df.iloc[:, 31], errors='coerce').fillna(0)
ae_data = pd.to_numeric(df.iloc[:, 30], errors='coerce').fillna(0)

total_sales = float(af_data.sum())
total_qty = int(ae_data.sum())
total_rows = int(len(df))

print(f'Total Sales: {total_sales}')
print(f'Total Qty: {total_qty}')
print(f'Total Rows: {total_rows}')

# 保存核心数据
result = {
    'overview': {
        'total_sales': total_sales,
        'total_sales_wan': round(total_sales / 10000, 2),
        'total_qty': total_qty,
        'total_rows': total_rows
    }
}

with open('data_core.json', 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False, indent=2)

print('Saved to data_core.json')
