# -*- coding: utf-8 -*-
"""
Created on Fri Jun 19 11:18:59 2020
@author: Yenny
"""

from faker import Faker
import random
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties

# 伪造数据
fake = Faker('zh_CN')  
name = []
score = []
age = []
number = range(1,21)
for _ in range(20):
    name.append(fake.name())
    score.append(random.randint(20,100))
    age.append(random.randint(18,20))

# 写入Excel
df = pd.DataFrame({
        '学号':number,
        '姓名':name,
        '年纪':age,
        '成绩':score
        })

df = df.set_index('学号')  
df.to_excel('学生成绩单.xlsx')

# 绘制图表
font = FontProperties(fname='C:\\Windows\\Fonts\\simfang.ttf', size=16)
students = pd.read_excel('学生成绩单.xlsx')
students.sort_values(by='成绩', inplace=True, ascending=False)
fig, ax=plt.subplots()
plt.bar(students.姓名, students.成绩, color='orange', edgecolor='none')
plt.title('学生成绩表', fontproperties=font, fontsize=16)
plt.xlabel('姓名', fontproperties=font, fontsize=14)
plt.ylabel('成绩', fontproperties=font, fontsize=14)
plt.xticks(students.姓名, fontproperties=font, rotation='90', fontsize=10)
plt.tight_layout()
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
plt.show()