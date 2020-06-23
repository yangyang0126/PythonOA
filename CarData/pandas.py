# -*- coding: utf-8 -*-
"""
Created on Tue Jun 23 12:09:36 2020
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
    table = pd.read_excel(path+name)
    table = table.replace([-40,0], np.nan)
    table.sort_values(by=['最高单体温度','最低单体温度'], inplace=True, ascending=[False,True], na_position='last')
    CarID.append(table.iloc[1,0])
    Time.append(table.iloc[0,1])
    TempMax.append(table.iloc[0,2])
    TempMin.append(table.iloc[0,3]) 
    table.sort_values(by='出水温度', inplace=True, na_position='last')
    if table.iloc[0,4] >= 20:
        Flag.append('不工作')
    else:
        Flag.append('水冷系统工作')
    table = table.dropna()
    Fault = table['热管理系统故障'].drop_duplicates().values.tolist()
    Fault = [str(int(x)) for x in Fault ]
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
    
