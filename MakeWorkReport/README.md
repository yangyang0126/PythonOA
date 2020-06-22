# 自动化生成周报

> GitHub：

## 流程

伪造数据，写入Excel

操作Excel，生成图表

写入Word文档

## 前期准备

需要用到的库

```python
from faker import Faker  # 伪造数据
import random  # 随机生成成绩
import pandas as pd  # 处理数据
import matplotlib.pyplot as plt  # 画数据图
from matplotlib.font_manager import FontProperties  # 设置图片的中文字体
from docx import Document  # 操作Word
from docx.oxml.ns import qn  # 设置Word字体
```

Faker：https://pypi.org/project/Faker/

pandas：https://pypi.org/project/pandas/

matplotlib：https://pypi.org/project/matplotlib/

python-docx：https://pypi.org/project/python-docx/

```bash
pip install Faker  
pip install pandas  # 用于操作数据
pip install matplotlib  # 用于数据可视化
pip install python-docx  # 操作Word
# `random`是Python自带的，不用安装
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

document = Document()
document.add_heading('学生成绩分析报告', level=0)
first_student = students.iloc[0,:]['姓名']
first_score = students.iloc[0,:]['总分']
# 设置格式
document.styles['Normal'].font.name = 'Times New Roman'
document.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')

p = document.add_paragraph('本次测评，全班共有{}名同学参加考试，其中分数总分排名第一的同学是'.format(len(students.姓名)))
p.add_run(str(first_student)).bold = True
p.add_run('，分数为')
p.add_run(str(first_score)).bold = True
p.add_run('。学生考试总体成绩如下')

table = document.add_table(rows=len(students.姓名)+1, cols=5, style='Medium Shading 1 Accent 5')
table.cell(0,0).text = '姓名'
table.cell(0,1).text = '语文'
table.cell(0,2).text = '数学'
table.cell(0,3).text = '英语'
table.cell(0,4).text = '总分'

for i,(index,row) in enumerate(students.iterrows()):
    table.cell(i+1, 0).text = str(row['姓名'])
    table.cell(i+1, 1).text = str(row['语文'])
    table.cell(i+1, 2).text = str(row['数学'])
    table.cell(i+1, 3).text = str(row['英语'])
    table.cell(i+1, 4).text = str(row['总分'])

document.add_picture('Part3_data.jpg')
document.save('Part3_学生成绩分析报告.docx')

```
document = Document()
document.add_paragraph
```

![](http://cdn.zhaojingyi0126.com/IMG/image-20200622155843178.png)

## 完整代码



```
from faker import Faker
import random
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
from docx import Document
from docx.oxml.ns import qn


# 伪造数据
fake = Faker('zh_CN')  
name = []
score1 = []
score2 = []
score3 = []
number = range(1,21)
for _ in range(20):
    name.append(fake.name())
    score1.append(random.randint(60,100))
    score2.append(random.randint(60,100))
    score3.append(random.randint(60,100))

# 写入Excel
df = pd.DataFrame({
        '学号':number,
        '姓名':name,        
        '语文':score1,
        '数学':score2,
        '英语':score3
        })

df = df.set_index('学号')  
df.to_excel('Part3_学生成绩单.xlsx')

# 读取数据
students = pd.read_excel('Part3_学生成绩单.xlsx')

# 排序名单
students['总分'] = students.语文 + students.数学 + students.英语
students.sort_values(by='总分', inplace=True, ascending=False)
students.reset_index(drop=True, inplace=True)

# 设置字体
font = FontProperties(fname='C:\\Windows\\Fonts\\simfang.ttf', size=16)

# 绘制图表
plt.rcParams['font.sans-serif']=['SimHei'] # 解决图例中文乱码
plt.rcParams['axes.unicode_minus']=False
ax = students.plot.bar(x='姓名', y=['语文','数学','英语'], stacked=True)
plt.title('学生成绩汇总图', fontproperties=font, fontsize=16)
plt.xlabel('姓名', fontproperties=font, fontsize=10)
plt.xticks(fontproperties=font, rotation='45', fontsize=8)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
plt.tight_layout()
plt.savefig('Part3_data.jpg')

# 操作Word
document = Document()
document.add_heading('学生成绩分析报告', level=0)
first_student = students.iloc[0,:]['姓名']
first_score = students.iloc[0,:]['总分']
# 设置格式
document.styles['Normal'].font.name = 'Times New Roman'
document.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')

p = document.add_paragraph('本次测评，全班共有{}名同学参加考试，其中分数总分排名第一的同学是'.format(len(students.姓名)))
p.add_run(str(first_student)).bold = True
p.add_run('，分数为')
p.add_run(str(first_score)).bold = True
p.add_run('。学生考试总体成绩如下')

table = document.add_table(rows=len(students.姓名)+1, cols=5, style='Medium Shading 1 Accent 5')
table.cell(0,0).text = '姓名'
table.cell(0,1).text = '语文'
table.cell(0,2).text = '数学'
table.cell(0,3).text = '英语'
table.cell(0,4).text = '总分'

for i,(index,row) in enumerate(students.iterrows()):
    table.cell(i+1, 0).text = str(row['姓名'])
    table.cell(i+1, 1).text = str(row['语文'])
    table.cell(i+1, 2).text = str(row['数学'])
    table.cell(i+1, 3).text = str(row['英语'])
    table.cell(i+1, 4).text = str(row['总分'])

document.add_picture('Part3_data.jpg')
document.save('Part3_学生成绩分析报告.docx')
```

