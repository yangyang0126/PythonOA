# -*- coding: utf-8 -*-
"""
Created on Tue Jun 30 14:18:59 2020
@author: Yenny
"""

import pandas as pd
from pyecharts import options as opts
from pyecharts.charts import Calendar

table = pd.read_excel('扇贝打卡统计.xls')
data = [[table.日期[i], int(table.学习时间[i])] for i in range(len(table))]

calendar=(
        Calendar()
        .add("", data, calendar_opts=opts.CalendarOpts(range_="2020"))
        .set_global_opts(
            title_opts=opts.TitleOpts(title="Shanbay"),
            visualmap_opts=opts.VisualMapOpts(
                max_=max(table.学习时间),
                min_=min(table.学习时间),
                orient="horizontal",
                is_piecewise=False,
                pos_top="230px",
                pos_left="100px",
#                range_color = ['#B84363','#F6E6EA'],  #自定义渐变颜色
            ),
        )
    )

calendar.render()