import pandas as pd
import numpy as np

M2_reate = pd.read_excel('首次M2逾期率/首次M2.xlsx')
data1 = pd.read_excel('人员清单.xlsx')
M2_reate = pd.merge(M2_reate,data1,on="SA工号",how="left")






