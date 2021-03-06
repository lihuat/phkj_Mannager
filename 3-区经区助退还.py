import numpy as np
import pandas as pd
import datetime

print("开始计算")
starttime = datetime.datetime.now()
m2_plus_back = pd.read_excel("M2+追回/M2+追回.xlsx",dtype={'SA工号':"O",'贷款编号':'O','SA工号':'O'})
group_all = pd.read_excel("M2+追回/区域风控跟踪表.xlsx",dtype={'SA工号':'O','贷款编号':'O','SA工号':'O','拿绩效区助工号':'O','拿绩效区经工号':'O'})



m2_plus_back1 = pd.merge(m2_plus_back,group_all,on="贷款编号",how="left",)
m2_plus_back1= m2_plus_back1[["贷款编号","SA工号_x","SA姓名_x","拿绩效区助工号","拿绩效区助姓名",
                              "拿绩效区经工号","拿绩效区经名称","扣押月","退还月份","区助扣罚金额",
                              "区经扣罚金额"]]
#转换数据类型
#m2_plus_back1[["扣押月",'SA工号_x','拿绩效区助工号',
#               '拿绩效区经工号']] = m2_plus_back1[["扣押月",'SA工号_x','拿绩效区助工号',
#              '拿绩效区经工号']].astype('O')


#判断是否已经在历史中扣押的单子
for i in range(len(m2_plus_back1)):
    if np.isnan(m2_plus_back1.loc[i,"扣押月"]):
        m2_plus_back1 = m2_plus_back1.drop(i)
    else:
        pass

#重置索引
m2_plus_back1 = m2_plus_back1.reset_index(drop=True)
#判断有无已经退还的单子
for i in range(len(m2_plus_back1)):
    if np.isnan(m2_plus_back1.loc[i,"退还月份"]) == False:
        m2_plus_back1 = m2_plus_back1.drop(i)
    else:
        pass

print("计算完成！正在保存文件...")
#数据输出保存为Excel格式的数据
m2_plus_back1.to_excel("输出/区域经理单笔退还.xlsx")
endtime = datetime.datetime.now()
print("用时：%d秒"%(endtime-starttime).seconds)

