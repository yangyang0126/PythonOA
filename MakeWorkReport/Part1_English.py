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
score = []
age = []
number = range(1,21)
for _ in range(20):
    name.append(fake.name())
    score.append(random.randint(20,100))
    age.append(random.randint(18,20))

# 写入Excel
df = pd.DataFrame({
        'Id':number,
        'Name':name,
        'Age':age,
        'Score':score
        })

df = df.set_index('Id')  
df.to_excel('students.xlsx')

# 绘制图表
students = pd.read_excel('students.xlsx')
students.sort_values(by='Score', inplace=True, ascending=False)
fig, ax=plt.subplots()
plt.bar(students.Name, students.Score, color='orange', edgecolor='none')
plt.title('Student Score', fontsize=16)
plt.xlabel('Name', fontsize=10)
plt.ylabel('Score', fontsize=10)
plt.xticks(students.Name, rotation='90', fontsize=8)
plt.tight_layout()
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
plt.show()