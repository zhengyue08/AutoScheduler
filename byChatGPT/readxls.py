import pandas as pd

# 读取 Excel 文件
filename = 'schedule.xlsx'
sheet_name = 'unavailable'
df = pd.read_excel(filename, sheet_name)

# 将 DataFrame 转换为字典
unavailable = {}
for index, row in df.iterrows():
    name = row['name']
    shifts = [x for x in row[1:] if not pd.isna(x)]
    unavailable[name] = shifts

print(unavailable)