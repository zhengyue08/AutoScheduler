import pandas as pd
import random

# 读取不可用时间段表格，生成不可用时间字典
unavailable_file = 'unavailable.xlsx'
unavailable_sheet = 'unavailable'
df = pd.read_excel(unavailable_file, sheet_name=unavailable_sheet)
unavailable_dict = {name: list(df.iloc[i, 1:].dropna()) for i, name in enumerate(df['name'])}

# 每个人的最大工作次数
MAX_WORK_TIMES = 2

# 一周工作日
WORK_DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']

# 每天的工作时间
WORK_TIMES = ['morning', 'noon', 'afternoon', 'night']

# 每个班次需要的人数
REQUIRED_PEOPLE = 2

# 所有工作时间段的列表
all_shifts = [(day, time) for day in WORK_DAYS for time in WORK_TIMES]

# 初始化结果字典，记录每个人的工作时间
schedule = {name: [] for name in unavailable_dict.keys()}

# 检查是否有足够的人来填充每个班次
def check_schedule():
    for shift in all_shifts:
        people = [name for name in unavailable_dict.keys() if shift not in schedule[name] and shift not in unavailable_dict[name]]
        if len(people) < REQUIRED_PEOPLE:
            return False
    return True

# 尝试给定的人员安排工作时间
def assign_work(name, shifts):
    for shift in shifts:
        if shift not in schedule[name] and shift not in unavailable_dict[name]:
            schedule[name].append(shift)
            return True
    return False

# 将工作时间转换为表格
def format_schedule():
    df = pd.DataFrame(schedule)
    df = df.transpose()
    df = df.applymap(lambda x: x[1] if len(x) > 0 else '')
    df.columns = WORK_DAYS
    return df

# 随机生成一周的工作安排
def generate_schedule():
    # 在所有可用人员中随机选择
    available_people = [name for name in unavailable_dict.keys() if len(schedule[name]) < MAX_WORK_TIMES]
    while check_schedule() and len(available_people) > 0:
        shift = random.choice(all_shifts)
        random.shuffle(available_people)
        assigned = False
        for name in available_people:
            if assign_work(name, [shift]):
                available_people.remove(name)
                assigned = True
                break
        if not assigned:
            break
    return format_schedule()

# 生成并输出工作安排
print(generate_schedule())
