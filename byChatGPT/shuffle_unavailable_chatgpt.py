import random
from openpyxl import Workbook

# 创建一个Workbook对象
wb = Workbook()

# 获取活跃的工作表对象
ws = wb.active

# 写入表头
ws.append(['name', 'weekday', 'period'])

# 随机生成一些不可用时间
names = ['Alice', 'Bob', 'Charlie', 'David', 'Eva', 'Lisa', 'Steve', 'Kobe', 'Richard', 'Rose', 'Keven', 'Pine', 'Rick', 'Sherlock', 'Jasmine', 'Alan', 'Cathy', 'Billy', 'Jane', 'Alex'] 
weekdays = ['星期一', '星期二', '星期三', '星期四', '星期五']
periods = ['12节', '34节', '56节', '78节']
unavailable = []
for i in range(80):
    name = random.choice(names)
    weekday = random.choice(weekdays)
    period = random.choice(periods)
    unavailable.append((name, weekday, period))

# 写入不可用时间到表格中
for row in unavailable:
    ws.append(row)

# 保存表格
wb.save('unavailable.xlsx')
