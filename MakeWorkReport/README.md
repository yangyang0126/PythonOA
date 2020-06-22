# 自动化生成周报

流程：

伪造数据，写入Excel

操作Excel，生成图表

写入Word文档

## 前期准备

需要用到的库

random：

Faker：https://pypi.org/project/Faker/

pandas：https://pypi.org/project/pandas/

matplotlib：https://pypi.org/project/matplotlib/

docx：https://pypi.org/project/docx/

```bash
pip install Faker  # 伪造数据
pip install pandas  # 用于操作数据
pip install matplotlib  # 用于数据可视化
pip install docx  # 操作Word
# `random`是Python自带的，不用安装
```

注意，在安装docx的时候，有可能会出现，你已经安装好了，但是依旧import失败，报错原因是`No module named 'exceptions'`，那就装一下`python-docx`

```
pip install python-docx
```

## 伪造数据

伪造数据，默认是英文的

如果你想构造中文数据，加入`zh_CN`，代表简体中文

```python
fake = Faker()  # 默认英文
fake = Faker('zh_CN')  # 如果需要简体中文，就改成这个
```

学号是1-20号

```python
number = range(1,21)
```

用Faker伪造姓名，成绩和年纪就用random在指定范围内随机产生

```python
name = []
score = []
age = []
for _ in range(20):
    name.append(fake.name())
    score.append(random.randint(20,100))
    age.append(random.randint(18,20))
```

## 写入Excel

构造数据

```python
df = pd.DataFrame({
        'Id':number,
        'Name':name,
        'Age':age,
        'Score':score
        })
```

将学号定义为索引

```python
df = df.set_index('Id')  
```

保存到Excel，不写路径的时候，默认和代码在一个文件夹下面

```python
df.to_excel('students.xlsx')
```

## 绘制图表

### 处理数据

取读Excel数据

```python
students = pd.read_excel('students.xlsx')
```

将数据根据`Score`（考试成绩）排序，`ascending=False`从大到小，`inplace`原地修改

```python
students.sort_values(by='Score', inplace=True, ascending=False)
```

### 绘制柱状图

根据名字和成绩绘制柱状图，颜色是橙色，不要边框

```python
plt.bar(students.Name, students.Score, color='orange', edgecolor='none')
```

设置坐标轴和标题的字体大小和内容

```python
plt.title('Student Score', fontsize=16)
plt.xlabel('Name', fontsize=10)
plt.ylabel('Score', fontsize=10)
plt.xticks(students.Name, rotation='90', fontsize=8)  # X轴，字体旋转90度
```

紧凑排布

```python
plt.tight_layout()
```

把图像边框去掉一部分（为了美观）

```python
fig, ax=plt.subplots()
--------
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
```

显示图像

```python
plt.show()
```

到这一步，完整代码是这样的

```python
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
```

这部分的完整代码可以在这里下载：

![](http://cdn.zhaojingyi0126.com/IMG/image-20200622111420913.png)

### 叠加柱状图

```python
# 计算三个月的总成绩就
students = pd.read_excel('Part2.xlsx')
students['Totel'] = students.Jan + students.Feb + students.Mar
students.sort_values(by='Totel', inplace=True)

# 绘制图表
students.plot.bar(x='Name', y=['Jan','Feb','Mar'], stacked=True)
plt.title('Student Score', fontsize=16)
plt.xlabel('Name', fontsize=10)
plt.tight_layout()
plt.show()
```

这部分的完整代码可以在这里下载：

![](http://cdn.zhaojingyi0126.com/IMG/image-20200622124056792.png)

### 中文支持

如果你想要一个中文效果的图，你需要一些额外的操作

导入中文支持

```python
from matplotlib.font_manager import FontProperties
```

设置字体

```python
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
```

这部分的完整代码可以在这里下载：

![](http://cdn.zhaojingyi0126.com/IMG/image-20200622121239281.png)

## 写入Word

```
document = Document()
document.add_paragraph
```

