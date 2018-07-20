import numpy as np
import pandas as pd

import numpy as np
import pandas as pd
#导入数据
M3 = pd.read_excel("首次M3明细/首次M3明细表.xlsx")
data1 = pd.read_excel("首次M3明细/绩效明细表.xlsx")
M3 = pd.merge(M3,data1,on="贷款编号",how="left")
people = pd.read_excel('首次M3明细/人员在职情况.xlsx')


for i in range(len(M3)):
    if np.isnan(M3.loc[i, "核算工号"]) == True:
        M3 = M3.drop(i)
    else:
        pass
M3 = M3.reset_index(drop=True)
M3['扣罚金额']=0

for i in range(len(M3)):
    if M3.loc[i,'产品名称'] == '003产品':
        M3.loc[i,'扣罚金额'] =20
    elif np.isnan(M3.loc[i,'区域经理助理编号'])==False:
        M3.loc[i,"扣罚金额"] = 18
    elif np.isnan(M3.loc[i,'区域经理编号'])==False:
        M3.loc[i,"扣罚金额"] = 30
    else:
        pass

M3 = pd.merge(M3,people,how="left",on="核算工号")

for i in range(len(M3)):
    if M3.loc[i,"在岗状态"] == "离职":
        M3 = M3.drop(i)
    elif pd.isnull(M3.loc[i,"在岗状态"]) == True:
        M3 = M3.drop(i)
    else:
        pass

M3.to_excel("输出/区经区助单笔扣罚.xls")



















