import random
import pandas as pd

# 定义人员姓名和时间段
names = ['Alice', 'Bob', 'Charlie', 'David', 'Eva', 'Lisa', 'Steve', 'Kobe', 'Richard', 'Rose', 'Keven', 'Pine', 'Rick', 'Sherlock', 'Jasmine', 'Alan', 'Cathy', 'Billy', 'Jane', 'Alex'] 
shifts = ["星期一 12节", "星期一 34节", "星期一 56节", "星期一 78节", 
          "星期二 12节", "星期二 34节", "星期二 56节", "星期二 78节",
          "星期三 12节", "星期三 34节", "星期三 56节", "星期三 78节",
          "星期四 12节", "星期四 34节", "星期四 56节", "星期四 78节",
          "星期五 12节", "星期五 34节", "星期五 56节", "星期五 78节"]

# 定义每个人在一周内已经值班的次数和时间段
person_shift_counts = {name: 0 for name in names}
person_shift_times = {name: [] for name in names}

# 读取不可用时间表格
unavailable_df = pd.read_excel("unavailable.xlsx")

# 将不可用时间转换为字典，方便后续查询
unavailable = {}
for _, row in unavailable_df.iterrows():
    name, shift, unavailabilities = row
    unavailable[(name, shift)] = unavailabilities
# print(unavailable)

# 分配班次
assigned = []
for shift in shifts:
    unassigned = names.copy()
    for i in range(2):
        # 获取可用的人员
        available = [name for name in unassigned if person_shift_counts[name] < 2 and shift not in person_shift_times[name] and name not in unavailable.get((name, shift), [])]
        if len(available) == 0:
            continue
        # 随机分配人员
        name = random.choice(available)
        person_shift_counts[name] += 1
        person_shift_times[name].append(shift)
        assigned.append((name, shift))
        unassigned.remove(name)
    # 如果只分配了一个人，则继续分配
    if len(assigned) % 2 != 0:
        name = random.choice(unassigned)
        person_shift_counts[name] += 1
        person_shift_times[name].append(shift)
        assigned.append((name, shift))

# 将结果保存为 Excel 文件
df = pd.DataFrame(assigned, columns=["Name", "Shift"])
df.to_excel("schedule.xlsx", index=False)

