import random

# 定义班次和人员名单
shifts = ['周一1', '周一2', '周一3', '周一4', '周二1', '周二2', '周二3', '周二4',
          '周三1', '周三2', '周三3', '周三4', '周四1', '周四2', '周四3', '周四4',
          '周五1', '周五2', '周五3', '周五4']
staff = ['张三', '李四', '王五', '赵六', '钱七', '孙八']

# 定义一个字典用于存储班次和对应的人员
schedule = {}

# 对于每个班次，从人员名单中随机选择两个人员，并将其分配到该班次
for shift in shifts:
    available_staff = list(staff)
    assigned_staff = []
    for i in range(2):
        if available_staff:
            selected_staff = random.choice(available_staff)
            assigned_staff.append(selected_staff)
            del available_staff[available_staff.index(selected_staff)]
        else:
            raise Exception('没有可用的人员')
    schedule[shift] = assigned_staff

# # 调整排班表以满足人员不可用的限制
# unavailable = {'张三': ['周二2', '周四1'], '李四': ['周一1', '周五4']}

# for name, shifts in unavailable.items():
#     for shift in shifts:
#         schedule[shift].remove(name)
#         available_staff = list(set(staff) - set(schedule[shift]))
#         if available_staff:
#             selected_staff = random.choice(available_staff)
#             schedule[shift].append(selected_staff)
#         else:
#             raise Exception('没有可用的人员')

# # 调整排班表以满足每人最多值两次班的限制
# for name in staff:
#     assigned_shifts = [shift for shift, people in schedule.items() if name in people]
#     if len(assigned_shifts) > 2:
#         for shift in assigned_shifts[2:]:
#             available_shifts = list(set(shifts) - set(assigned_shifts[:2]))
#             if available_shifts:
#                 selected_shift = random.choice(available_shifts)
#                 schedule[shift].remove(name)
#                 schedule[selected_shift].append(name)
#                 assigned_shifts = [s for s in assigned_shifts if s != shift] + [selected_shift]
#             else:
#                 raise Exception('没有可用的班次')

# 打印生成的班次表
for shift, people in schedule.items():
    print(shift, ':', people)
