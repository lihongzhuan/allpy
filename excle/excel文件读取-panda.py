# -*- coding: utf-8 -*-
"""
Created on Fri Jun 19 22:13:49 2020

@author: Administrator
"""


import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime
from pandas import Series,DataFrame

df = pd.read_excel('d:/RIC组装表.xlsx') #读取xlsx文件
df.shape
df.dtypes

df.columns

df.tail(5)
print("完成")

