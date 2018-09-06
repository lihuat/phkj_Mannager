
"""

1、绩效明细表：前推三个月和前推两个月的绩效数据
2、销售人员清单：当月31日/30日的
3、首次M3明细表：要剔除不在当期的单子，SA姓名下去掉“虚拟”，“服务”，“通讯”。大区只保留【北，南，西，直管】，
产品名称剔除“产品1，产品2...”

"""

print("请先注意检查以下项：\n1、绩效明细表：前推三个月和前推两个月的绩效数据")
print("2、销售人员清单：当月31日/30日的")
print("3、首次M3明细表：要剔除不在当期的单子，SA姓名下去掉“虚拟”，“服务”，“通讯”。大区只保留【北，南，西，直管】，产品名称剔除“产品1，产品2...”")
print("4、请先计算区经区助扣罚免除")

import pandas as pd
from tqdm import tqdm
import numpy as np
import datetime
print("开始计算")
starttime = datetime.datetime.now()
#导入数据
M3 = pd.read_excel("首次M3明细/首次M3明细表.xlsx",dtype={'贷款编号':'O','SA工号':'O'})
data1 = pd.read_excel("首次M3明细/绩效明细表.xlsx",dtype={'贷款编号':'O','核算工号':'O','区域经理助理编号':'O','区域经理编号':'O'})
M3 = pd.merge(M3,data1,on="贷款编号",how="left")
#导入销售人员清单
people = pd.read_excel('首次M3明细/销售人员清单.xlsx',dtype={'工号':'O'})
people = people[['工号','在岗状态']]
#导入区经区助免除扣罚名单
Free = pd.read_excel("输出/区经区助扣罚免除.xlsx",dtype={"区经工号":'O'})
Free = Free[["区经工号","是否免扣"]]

#选取字段
M3=M3[['贷款编号','SA工号','SA姓名','产品名称','商户','门店','区域经理助理编号','区域经理助理姓名','区域经理编号','区域经理姓名']]

#区域经理助理扣罚
M3_qz = M3.copy()
#提取区助
for i in tqdm(range(len(M3_qz))):
    if pd.isnull(M3_qz.loc[i,"区域经理助理编号"]) ==True:
        M3_qz = M3_qz.drop(i)
    else:
        pass
M3_qz = M3_qz.reset_index(drop=True)#重置索引

people_qz = people.copy()#复制一份销售人员清单的在职状态
people_qz = people_qz.rename(index=str,columns={"工号":'区域经理助理编号'})
M3_qz = pd.merge(M3_qz,people_qz,on="区域经理助理编号",how="left")
M3_qz = M3_qz[M3_qz["在岗状态"]=="在职"]#区助在职的人员
M3_qz = M3_qz.reset_index(drop=True)#重置索引
#扣罚逻辑
for i in range(len(M3_qz)):
    if M3_qz.loc[i,"产品名称"] == '003产品':
        M3_qz.loc[i,"区助扣罚金额"] = 20
    else:
        M3_qz.loc[i,"区助扣罚金额"] = 18
M3_qz["区经扣罚金额"] = ""
#区助是否免扣
Free_qz = Free.copy()
Free_qz = Free_qz.rename(index=str,columns={"区经工号":'区域经理助理编号'})
M3_qz = pd.merge(M3_qz,Free_qz,on="区域经理助理编号",how="left")
M3_qz = M3_qz[M3_qz["是否免扣"] =="扣罚"]

#区域经理首次M3扣罚
M3_qj = M3.copy()#复制一份
people_qj = people.copy()#复制一份销售人员清单在职状态
people_qj = people_qj.rename(index=str,columns={"工号":'区域经理编号'})#重命名
for i in tqdm(range(len(M3_qj))):
    if pd.isnull(M3_qj.loc[i,"区域经理编号"]) ==True:
        M3_qj = M3_qj.drop(i)
    else:
        pass
M3_qj = M3_qj.reset_index(drop=True)
M3_qj = pd.merge(M3_qj,people_qj,on="区域经理编号",how="left")
M3_qj = M3_qj[M3_qj["在岗状态"]=="在职"]
M3_qj = M3_qj.reset_index(drop=True)

M3_qj['区助扣罚金额'] = ""

for i in range(len(M3_qj)):
    if M3_qj.loc[i,"产品名称"] == '003产品':
        M3_qj.loc[i,"区经扣罚金额"] = 20
    else:
        M3_qj.loc[i,"区经扣罚金额"] = 30

#区经是否免扣
Free_qj = Free.copy()
Free_qj = Free_qj.rename(index=str,columns={"区经工号":'区域经理编号'})
M3_qj = pd.merge(M3_qj,Free_qj,on="区域经理编号",how="left")
M3_qj = M3_qj[M3_qj["是否免扣"] =="扣罚"]

#合并区经和区助的扣罚
print("计算完成！正在保存文件")
M3_result = pd.concat([M3_qj,M3_qz],sort=False)
M3_result.to_excel("输出/区经区助单笔扣罚.xlsx")
print("文件保存成功！good luck!")
endtime = datetime.datetime.now()
print("用时：%d秒"%(endtime-starttime).seconds)