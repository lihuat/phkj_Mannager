
import pandas as pd
from tqdm import tqdm
#导入数据
M3 = pd.read_excel("首次M3明细/首次M3明细表.xlsx",dtype={'贷款编号':'O','SA工号':'O'})
data1 = pd.read_excel("首次M3明细/绩效明细表.xlsx",dtype={'贷款编号':'O','核算工号':'O','区域经理助理编号':'O','区域经理编号':'O'})
M3 = pd.merge(M3,data1,on="贷款编号",how="left")
M3=M3[['贷款编号','SA工号','SA姓名','产品名称','核算工号','核算姓名','区域经理助理编号','区域经理编号','高区名称','大区名称']]

for i in tqdm(range(len(M3))):
    if pd.isnull(M3.loc[i,"核算工号"]) == True:
        M3 = M3.drop(i)
    else:
        pass
M3 = M3.reset_index(drop=True)
M3['扣罚金额']=0

for i in tqdm(range(len(M3))):
    if M3.loc[i,'产品名称'] == '003产品':
        M3.loc[i,'扣罚金额'] =20
    elif pd.isnull(M3.loc[i,'区域经理助理编号'])==False:
        M3.loc[i,"扣罚金额"] = 18
    elif pd.isnull(M3.loc[i,'区域经理编号'])==False:
        M3.loc[i,"扣罚金额"] = 30
    else:
        pass

people = pd.read_excel('首次M3明细/销售人员清单.xlsx',dtype={'工号':'O'})
people = people[['工号','在岗状态']]
people = people.rename(index=str,columns={"工号":'核算工号'})

M3 = pd.merge(M3,people,how="left",on="核算工号")

for i in tqdm(range(len(M3))):
    if M3.loc[i,"在岗状态"] == "离职":
        M3 = M3.drop(i)
    elif pd.isnull(M3.loc[i,"在岗状态"]) == True:
        M3 = M3.drop(i)
    else:
        pass

M3 = M3[['贷款编号','SA工号','SA姓名','产品名称','核算工号','核算姓名','在岗状态','高区名称','大区名称','扣罚金额']]
M3.to_excel("输出/区经区助单笔扣罚.xlsx")
