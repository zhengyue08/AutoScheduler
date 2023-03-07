import random
import pandas as pd
import openpyxl

# define names and shifts
names = ['张三 Richard', 
         '李四 Rick', 
         '王二 Sherlock', 
         '唐三 Jane', 
         '唐四 Jasmine', 
         '唐五 Alan', 
         '王三 Keith', 
         '李三 Jane', 
         '李四 Sydnie', 
         '张三 Julius', 
         '李四 Anna', 
         '王二 Alvin', 
         '唐三 Alex', 
         '唐四 Pine', 
         '唐五 Cathy', 
         '王三 Juliet', 
         '李三 Jane'] 

shifts = ["星期一 12节", "星期一 34节", "星期一 56节", "星期一 78节", 
          "星期二 12节", "星期二 34节", "星期二 56节", "星期二 78节",
          "星期三 12节", "星期三 34节", "星期三 56节", "星期三 78节",
          "星期四 12节", "星期四 34节", "星期四 56节", "星期四 78节",
          "星期五 12节", "星期五 34节", "星期五 56节", "星期五 78节"]

# Interface：Input 
instructions = eval(input("请输入你想要的操作：\n1. 生成值班表"))

if instructions == 1:
    week = eval(input("您需要为第几周排班？"))
elif instructions == 2:
    add_name = input("请输入学工助理的姓名和英文名（格式：小明 Xiao Ming）：")
    names.append(add_name)
elif instructions == 3:
    delete_name = input("请输入学工助理的姓名和英文名（格式：小明 Xiao Ming）：")
    names.remove(delete_name)

def write_to_excel (assigned_by_shift):
    # 将结果保存为 Excel 文件
    df = pd.DataFrame(assigned_by_shift, columns=["Name", "Shift"])
    df.to_excel("schedule.xlsx", index=False)


def drawExcel(assigned_by_shift, week):

    # 创建一个Workbook对象
    wb = openpyxl.load_workbook('templeUP.xlsx')

    # 获取活跃的工作表对象
    ws = wb.active

    # 写入表头
    ws['A1'] = "第" + str(week) + "周值班表"

    ws['A2'] = ""
    ws['A3'] = "12节"
    ws['A4'] = "34节"
    ws['A5'] = "56节"
    ws['A6'] = "78节"

    ws['B2'] = "星期一"
    ws['C2'] = "星期二"
    ws['D2'] = "星期三"
    ws['E2'] = "星期四"
    ws['F2'] = "星期五"

    ws['B3'] = ' '.join(assigned_by_shift['星期一 12节'])
    ws['B4'] = ' '.join(assigned_by_shift['星期一 34节'])
    ws['B5'] = ' '.join(assigned_by_shift['星期一 56节'])
    ws['B6'] = ' '.join(assigned_by_shift['星期一 78节'])


    ws['C3'] = ' '.join(assigned_by_shift['星期二 12节'])
    ws['C4'] = ' '.join(assigned_by_shift['星期二 34节'])
    ws['C5'] = ' '.join(assigned_by_shift['星期二 56节'])
    ws['C6'] = ' '.join(assigned_by_shift['星期二 78节'])

    ws['D3'] = ' '.join(assigned_by_shift['星期三 12节'])
    ws['D4'] = ' '.join(assigned_by_shift['星期三 34节'])
    ws['D5'] = ' '.join(assigned_by_shift['星期三 56节'])
    ws['D6'] = ' '.join(assigned_by_shift['星期三 78节'])

    ws['E3'] = ' '.join(assigned_by_shift['星期四 12节'])
    ws['E4'] = ' '.join(assigned_by_shift['星期四 34节'])
    ws['E5'] = ' '.join(assigned_by_shift['星期四 56节'])
    ws['E6'] = ' '.join(assigned_by_shift['星期四 78节'])

    ws['F3'] = ' '.join(assigned_by_shift['星期五 12节'])
    ws['F4'] = ' '.join(assigned_by_shift['星期五 34节'])
    ws['F5'] = ' '.join(assigned_by_shift['星期五 56节'])
    ws['F6'] = ' '.join(assigned_by_shift['星期五 78节'])

    for cell in range(8,8+len(names)):
        ws["D"+str(cell)] = names[cell-8]
        ws["E"+str(cell)] = person_shift_counts[names[cell-8]]

    wb.save("第" + str(week) + "周值班表最终版.xlsx")




# define two lists to store the number of shifts assigned for each person and the shifts of each person
person_shift_counts = {name: 0 for name in names}
person_shift_times = {name: [] for name in names}

# print(person_shift_counts)
# print(person_shift_times)

# high_priority -> two people per shift
high_priority_shifts = ["星期二 34节", "星期二 56节", "星期四 34节", "星期四 56节"]
normal_shifts = [shift for shift in shifts if shift not in high_priority_shifts]
print(normal_shifts)
# read unavailable periods from a excel file
unavailable_df = pd.read_excel("./unavailableUP.xlsx").dropna(axis=0, how='all')

unavailable = {}
available = {}

# store unavailable periods in a dictionary unavailable = e.g. "星期一 12节":["张稳 Rick", "辛晨 Sherlock"]
for _, row in unavailable_df.iterrows():
    name, day, period = row
    # print(name, shift, unavailabilities)
    shift = day + " " + period
    unavailable.setdefault(shift,[]).append(name)

# store available p1eriods in a dictionary: available = e.g. "星期一 12节":["郑越 Richard", "陶婧怡 Jane", "曾洁 Jasmine", "龚忠信 Alan", "高子轶 Keith", "胡倩雯 Jane", "金昕怡 Sydnie", "刘夏松 Julius", "李星星 Anna", "孙飞帅 Alvin", "吴金隆 Alex", "韦彦昊 Pine", "杨赟雪 Cathy", "张邱一果 Billy J", "朱丽叶 Juliet"]
for shift in unavailable:
    available[shift] = [name for name in names if name not in unavailable[shift]]


# store assigned shifts 
assigned_by_person = {}
assigned_by_shift = {}

# Prioritize high-priority shifts for assignment
for shift in high_priority_shifts:
    unassigned = names.copy()

    # assign two people per shift
    for i in range(2):
        # get available people 
        available_for_this_shift = [name for name in unassigned if person_shift_counts[name] < 2 and shift not in person_shift_times[name] and name in available.get((shift), [])]
        if len(available_for_this_shift) == 0:
            print("No available person for shift: ", shift)
            continue
        # assign person randomly
        name = random.choice(available_for_this_shift)

        # update available according to conditions
        person_shift_counts[name] += 1
        person_shift_times[name].append(shift)

        assigned_by_person.setdefault(name,[]).append(shift)
        assigned_by_shift.setdefault(shift,[]).append(name)

        # avoid assigning the same person to the same shift twice
        unassigned.remove(name)

# for shift in high_priority_shifts:
#     print(shift, assigned_by_shift[shift])
#     print(shift, available[shift])

# assign remaining shifts: make sure that every shift has at least one person assigned
for shift in normal_shifts:
    unassigned = names.copy()
    for i in range(1):
        # get available people 
        available_for_this_shift = [name for name in unassigned if person_shift_counts[name] < 2 and shift not in person_shift_times[name] and name in available.get((shift), [])]
        if len(available_for_this_shift) == 0:
            print("No available person for shift under the condition that each person is assigned to at most two shifts, try to assign someone for three shifts in this week", shift)
            available_for_this_shift = [name for name in unassigned if person_shift_counts[name] < 3 and shift not in person_shift_times[name] and name in available.get((shift), [])]

        # assign person randomly
        name = random.choice(available_for_this_shift)

        # update available according to conditions
        person_shift_counts[name] += 1
        person_shift_times[name].append(shift)

        assigned_by_person.setdefault(name,[]).append(shift)
        assigned_by_shift.setdefault(shift,[]).append(name)

        # avoid assigning the same person to the same shift twice
        unassigned.remove(name)

# assign one more shift to person who has only one shift assigned
one_shift_person = [name for name in names if person_shift_counts[name] == 1]
one_person_shift = [shift for shift in shifts if len(assigned_by_shift[shift]) == 1]

for name in one_shift_person:
    for shift in one_person_shift:
        if name in available.get((shift), []):
            person_shift_counts[name] += 1
            person_shift_times[name].append(shift)

            assigned_by_person.setdefault(name,[]).append(shift)
            assigned_by_shift.setdefault(shift,[]).append(name)

            # avoid assigning the same person to the same shift twice
            one_person_shift.remove(shift)
            break

print(len(assigned_by_shift))
print(assigned_by_shift)
print(person_shift_counts)

write_to_excel(assigned_by_shift)

drawExcel(assigned_by_shift, week)

