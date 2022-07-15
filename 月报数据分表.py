import pandas as pd
import numpy as np
from IPython.display import display
import warnings
from openpyxl import load_workbook
warnings.filterwarnings("ignore")

ym = '2022.06'
path = "D:\\1.office_syx\\不良\\2022.06不良月报\\"

#
# 加载excel
wb=pd.read_excel(path+ym+"不良明细终.xlsx",sheet_name=None)#读取excel所有sheet数据
# 切换到第一张表---明细全表
bs = wb['不良明细总表']
# print(bs)

# print(type(bs))
# 切换到第二张表---错漏明细表
bs_cl = wb['错漏明细']
#
br = pd.read_excel(path+ym+'不良率终.xlsx')
#
list_code = (100105,100118,100120,100122,100140,100141,103195,103233,103275,103407,103695,103730,104657,104738,104656,105042,105041,105078,105171,105630,105730,105929)
#
#
br.to_excel("000.xls")

# print(br)
for p in range(0,len(list_code)):
    aaa = br[br["供应商"] == list_code[p]]
    aaa['实际总不良率']= aaa['实际总不良率'].astype(float)
    aaa['实际总不良率'] = aaa['实际总不良率'].apply(lambda x: format((x+0.000001), '.2%'))
    # aaa.iloc[aaa.shape[0] - 1, 6] = format(aaa.iloc[aaa.shape[0] - 1, 6], '.2%')
    # aaa.iloc[aaa.shape[0] - 1, 8] = format(aaa.iloc[aaa.shape[0] - 1, 8], '.2%')
    # aaa.iloc[aaa.shape[0] - 1,10] = format(aaa.iloc[aaa.shape[0] - 1,10], '.2%')
#
    display(aaa)
    bbb = bs[bs["供应商"] == list_code[p]]
    ccc = bs_cl[bs_cl["供应商"] == list_code[p]]
    with pd.ExcelWriter(r'D:/1.office_syx/月报汇报/{}.xlsx'.format(list_code[p])) as writer:
        aaa.to_excel(writer,sheet_name='不良率',index = False)
        bbb.to_excel(writer,sheet_name='不良明细',index = False)
        ccc.to_excel(writer,sheet_name='错漏',index = False)