# -*- coding: utf-8 -*-
"""
Created on Fri Jun 19 11:18:59 2020
@author: Yenny
"""

from faker import Faker
import random
import pandas as pd
import matplotlib.pyplot as plt

# 伪造数据
fake = Faker()  
name = []
score1 = []
score2 = []
score3 = []
number = range(1,21)
for _ in range(20):
    name.append(fake.name())
    score1.append(random.randint(10,30))
    score2.append(random.randint(10,30))
    score3.append(random.randint(10,40))

# 写入Excel
df = pd.DataFrame({
        'Id':number,
        'Name':name,
        'Jan':score1,
        'Feb':score2,
        'Mar':score3
        })

df = df.set_index('Id')  
df.to_excel('Part2.xlsx')

# 计算三个月的总成绩就
students = pd.read_excel('Part2.xlsx')
students['Totel'] = students.Jan + students.Feb + students.Mar
students.sort_values(by='Totel', inplace=True, ascending=False)

# 绘制图表
ax = students.plot.bar(x='Name', y=['Jan','Feb','Mar'], stacked=True)
plt.title('Student Score', fontsize=16)
plt.xlabel('Name', fontsize=10)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
plt.tight_layout()
plt.show()