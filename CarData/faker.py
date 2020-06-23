# -*- coding: utf-8 -*-
"""
Created on Tue Jun 23 09:55:05 2020
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
        
    # 随机加入扰动数据
    for _ in range(5):
        i = random.randint(1, NumData-1)
        TempMax[i] = TempMin[i] = TempWater[i] = random.choice([0,-40])        
        j = random.randint(1, NumData-1)
        TempMax[j] = TempMin[j] = TempWater[j] = FaultID[j] = ''   
   
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