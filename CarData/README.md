# 汇总100个Excel表

> 甲方爸爸的需求：我有100台车的数据（100个EXCLE表），我需要从每个表里找到【出水温度】最小值（还需要提前判断-40，0这两个数值，-40和0不参与统计），并判断是否小于20度，小于则输出“水冷系统工作”，否则输出“不工作”。 还需要找到【单体最高温度】最大值，同时输出这一行【时间点】【单体最低温度】【单体最高温度】三项 

这个表格长这样（数据是自己伪造的）

![](http://cdn.zhaojingyi0126.com/IMG/image-20200623134729373.png)

然后这样的表格有100个

![](http://cdn.zhaojingyi0126.com/IMG/image-20200623134831145.png)

需要汇总的结果长这样

![](http://cdn.zhaojingyi0126.com/IMG/image-20200623135015267.png)

其实整个代码很简单，我反而是在伪造数据上花了比较多的时间。。因为甲方爸爸不能给我数据，我还得各种伪造，加入特殊值、异常值等……

## 伪造数据

```python
# -*- coding: utf-8 -*-
"""
Created on Tue Jun 23 2020
@author: Yenny
"""

from faker import Faker
import random
import pandas as pd
import os
 
def CreateData(NumTable, NumData, WorkFlag):
    CarID = []
    Time = pd.date_range('2020-06-08', periods=NumData, freq='20min')
    TempMax = []
    TempMin = []
    TempWater = []
    FaultID = []
    fake = Faker()  
    CarID.append(fake.password(length=15, digits=True, upper_case=True, lower_case=False, special_chars=False))
    CarID = [val for val in CarID for i in range(NumData)]
    for i in range(NumData):     
        if WorkFlag == 0:           
            TempMax.append(random.randint(20, 35))
            TempMin.append(TempMax[i]-random.randint(1, 2))
            TempWater.append(random.randint(20, 35))
            FaultID.append(0)
        else:
            TempMax.append(random.randint(5, 19))
            TempMin.append(TempMax[i]-random.randint(1, 2))
            TempWater.append(random.randint(5, 19))
            FaultID.append(0)
        
    # 随机加入干扰数据
    for _ in range(5):
        i = random.randint(1, NumData-1)
        TempMax[i] = TempMin[i] = TempWater[i] = random.choice([0,-40])        
        j = random.randint(1, NumData-1)
        TempMax[j] = TempMin[j] = TempWater[j] = FaultID[j] = ''   
    # 随机加点故障码
    n = random.randint(1, 3)
    for _ in range(n):
        k = random.randint(1, NumData-1)
        FaultID[k] = random.choice([0,68,38,101,555])      
            
    # 写入Excel
    df = pd.DataFrame({
            '车辆Vin':CarID,
            '时间':Time,
            '最高单体温度':TempMax,
            '最低单体温度':TempMin,
            '出水温度':TempWater,
            '热管理系统故障':FaultID
            })
    
    df = df.set_index('车辆Vin')  
    path = os.getcwd()
    path = path + '\\TotelData\\'
    TitleExcel = str(NumTable)
    df.to_excel(path + 'Data' + TitleExcel + '.xlsx')

# 伪造数据，100个表格，每个表格1000行
NumData = 1000
for NumTable in range(100):
    WorkFlag = random.choice([0, 1])
    CreateData(NumTable, NumData, WorkFlag)
```

## 汇总数据

在上一步，我们把伪造的100张表格，存到了一个叫`TotelData` 的文件夹里面，这一步就是读取里面的数据，进行汇总

![](http://cdn.zhaojingyi0126.com/IMG/image-20200623135641042.png)

```python
# -*- coding: utf-8 -*-
"""
Created on Tue Jun 23 2020
@author: Yenny
"""

import pandas as pd
import numpy as np
import os

path = os.getcwd()+ '\\TotelData\\'
PathList = os.listdir(path)

CarID = []
Time = []
TempMax = []
TempMin = []
FaultID = []
Flag = [] 

for name in PathList:
    
    # 依次读取表格
    table = pd.read_excel(path+name)
    
    # 先把-40和0都替换成空值
    table = table.replace([-40,0], np.nan)
    
    # 表格排序，获取“最高单体温度”最大的那一行
    # 如果有重复的行，再根据“最低单体温度”最小值筛选
    table.sort_values(by=['最高单体温度','最低单体温度'], inplace=True, ascending=[False,True], na_position='last')
    
    # 赋值
    CarID.append(table.iloc[1,0])
    Time.append(table.iloc[0,1])
    TempMax.append(table.iloc[0,2])
    TempMin.append(table.iloc[0,3]) 
    
    # 再根据“出水温度”排序
    table.sort_values(by='出水温度', inplace=True, na_position='last')
    if table.iloc[0,4] >= 20:
        Flag.append('不工作')
    else:
        Flag.append('水冷系统工作')
    table = table.dropna()
    
    # 把“系统故障”这一栏，非重复的值筛选出来
    Fault = table['热管理系统故障'].drop_duplicates().values.tolist()
    Fault = [str(int(x)) for x in Fault ]
    # 转成字符串格式
    Fault = "、".join(Fault)
    FaultID.append(Fault)

# 写入Excel
count = list(range(1,len(PathList)+1))  
df = pd.DataFrame({
        '序号':count,
        'VIN码':CarID,
        '单体最高温度':TempMax,
        '单体最低温度':TempMin,
        '最高温度时间点':Time,
        '水冷是否工作':Flag,
        '热管理系统故障':FaultID
        })
df = df.set_index('序号')     
df.to_excel(os.getcwd() + '\\汇总表.xlsx')  
```

## 其他

代码可在以下链接处下载：https://github.com/yangyang0126/PythonOA/tree/master/CarData

请先运行`faker'py`，这个代码伪造了100个Excel的数据

再运行`pandas.py`，这个代码形成汇总表