import pandas as pd
import numpy as np
from tqdm import tqdm
import re
import datetime

"""


"""

print("注意事项：\n销售人员清单：当月30日/31日")
print("计算开始")
starttime = datetime.datetime.now()

m2 = pd.read_excel("首次M2逾期率/首次M2逾期率.xlsx",dtype={"SA工号":'O'})
m2 = m2[['SA工号','SA姓名','首次M2单量','首次M2注册数']]

people = pd.read_excel("首次M3明细/销售人员清单.xlsx",dtype={'工号':'O'})
people = people[['工号','姓名','在岗状态','区域经理','高区经理','城市经理','角色']]
people = people.rename(columns={'工号':'SA工号'})

m2 = pd.merge(m2,people,on="SA工号",how="left",)
m2 = m2.loc[m2["在岗状态"]=="在职"]

dt1 = m2.loc[m2["角色"]=="销售员" ]
dt2 = m2.loc[m2["角色"]=="APP功能试点,销售员"]

m2 = pd.concat([dt1,dt2]).reset_index(drop=True)


m2['区经姓名'] = 0
for i in tqdm(range(len(m2))):
    try:
            m2.loc[i,'区经姓名'] = re.sub(r"团队$","",m2.loc[i,"区域经理"])
    except:
            pass

for i in tqdm(range(len(m2))):
     try:
            m2.loc[i,'区经姓名']= re.findall(r'[^（）]+',m2.loc[i,'区域经理'])[1]
     except:
            pass

m2['匹配'] = 0
for i in tqdm(range(len(m2))):
    m2.loc[i,'匹配'] = m2.loc[i,'高区经理']+m2.loc[i,'区经姓名']

people_1 = people.copy()

a1 = people_1.loc[people_1['角色']== "区域经理助理"]
a2 = people_1.loc[people_1['角色']== "区域经理"]
a3 = people_1.loc[people_1['角色']== "APP功能试点,区域经理助理"]
a4 = people_1.loc[people_1['角色']== "APP功能试点,区域经理"]

people_1 = pd.concat([a1,a2,a3,a4])
people_1 = people_1.reset_index(drop=True)

people_1['匹配'] = ""

for i in tqdm(range(len(people_1))):
    try:
        people_1.loc[i,'匹配'] = people_1.loc[i,'高区经理']+ people_1.loc[i,'姓名']
    except:
        pass

people_1 = people_1[["匹配",'SA工号']]


people_1 = people_1.rename(columns={'SA工号':"区经工号"})
m2 = pd.merge(m2,people_1,on="匹配",how="left")

df1 = m2.groupby(["城市经理","区经工号","区经姓名"],as_index=False)["首次M2单量","首次M2注册数"].sum()

df1["首次M2逾期率"] = 0
for i in tqdm(range(len(df1))):
    df1.loc[i, "首次M2逾期率"] = df1.loc[i, "首次M2单量"] / df1.loc[i, "首次M2注册数"]

df1["排名"] = df1["首次M2逾期率"].groupby(df1["城市经理"]).rank(ascending=1,method='first')
count_of_daqu =  df1.groupby(["城市经理"],as_index=False)["区经工号"].count()

df1 = pd.merge(df1,count_of_daqu,on="城市经理",how="left")

for i in range(len(df1)):
    if df1.loc[i,"排名"] <= round(df1.loc[i,"区经工号_y"]*0.6):
        df1.loc[i,"是否免扣"] = "免扣"
    else:
        df1.loc[i,"是否免扣"] = "扣罚"

df1 = df1.rename(index=str,columns = {'区经工号_x':'区经工号'})
print("计算完成，正在保存文件...")
df1.to_excel("输出/区经区助扣罚免除.xlsx")
endtime = datetime.datetime.now()
print("用时：%d秒"%(endtime-starttime).seconds)