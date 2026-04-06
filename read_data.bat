@echo off
chcp 65001 >nul
echo 读取机械革命销售日报数据...
python -c "import pandas as pd; import sys; df = pd.read_excel(r'C:\Users\刘莎\Desktop\机械革命 - 销售日报 3 月.xlsx'); print('列名:', list(df.columns)); print('行数:', len(df))"
pause
