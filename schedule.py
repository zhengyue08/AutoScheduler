import random
import pandas as pd
import openpyxl
import csv
from schedule_function import *


# unchangable values for every week and for any task

days = ["星期一", "星期二", "星期三", "星期四", "星期五",]

periods = ["12节", "34节", "56节", "78节"]

shifts = [day + " " + period for day in days for period in periods]



# Interface：Input 
instructions = eval(input("请输入你想要的操作：\n0. 确认本周值班人员和人数（强烈推荐每周排班前确认）\n1. 根据available_xth.xlsx排班\n2. 添加学工助理\n3. 删除学工助理\n4. 获取未填写的available_x.xlsx\n"))

if instructions == 0:
    print(get_names())
    print("本周值班人员共有" + str(len(get_names())) + "人")
elif instructions == 1:
    add_name = input("请输入学工助理的姓名和英文名（格式：小明 Xiao Ming）：")
    add_member(add_name)
elif instructions == 2:
    delete_name = input("请输入学工助理的姓名和英文名（格式：小明 Xiao Ming）：")
    delete_member(delete_name)
elif instructions == 3:
    week = eval(input("您需要为第几周生成未填写的available_x.xlsx？"))
    generate_temple_available(week)
elif instructions == 4:
    names = get_names()
    week = eval(input("您需要为第几周排班？"))
    generate_schedule(week, names)













