#导入包-------------------------------------------------------------
import pandas as pd
from IPython.display import display
import warnings,re
import datetime,time
from datetime import timedelta
from sqlalchemy import create_engine
import xlwt,xlrd
from xlwt import easyxf,Workbook
warnings.filterwarnings("ignore")

#1、时间定义-------------------------------------------------------------
today = time.strftime("%Y-%m-%d")  #2022-04-19
yesterday = datetime.datetime.today()+timedelta(-1)
yesterday_format = '2022.06.12'#yesterday.strftime('%Y.%m.%d')
month = yesterday.month
yesterday_format11 = '2022-06-12'#yesterday.strftime('%Y-%m-%d')

# 文件保存路径
file_path = "D:\\1.office_syx\\不良\\2022." + str(month) + "不良\\"

#数据库引擎
engine_detail_all= create_engine('mysql+pymysql://root:DR003812@localhost:3306/detail_all?charset=utf8')
engine_rate_all= create_engine('mysql+pymysql://root:DR003812@localhost:3306/rate_all?charset=utf8')

#从库detail_all中导出105表格数据
sql = 'select * from detail_105'
YIELD = pd.read_sql_query(sql=sql, con=engine_detail_all)
YIELD = YIELD.drop_duplicates()  #作去重处理
YIELD.to_sql(name='detail_105',con=engine_detail_all,if_exists='replace',index=False,index_label=False)    #放回数据库

#导出累计105表格
leiji = pd.read_excel(file_path+'累计-105.xlsx')
leiji_quality = leiji[leiji['责任归属'].isin(['供应商责任'])]  # 筛选出品质差错

ls = ['/漏刻字', '/刻字错', '/石重低配', '/镶错石', '/款式错', '/材质错', '/手寸错', '/成色不足']

new= pd.DataFrame(columns = ['年度','月份','供应商','PO','PO项目','物料凭证','物料凭证年度','物料凭证项目','款号','批次','质检序号','质检状态','不良原因','责任归属','数量','移动类型','过账日期','采购订单类型','定制编码'
])

for j in range(0,leiji_quality.shape[0]):
    for i in ls:
        if i in str(leiji_quality.iloc[j,12]):
            TTT = pd.DataFrame(leiji_quality.loc[j]).T
            new = new.append(TTT,ignore_index=True)
display(new)

#从库detail_all中导出品质差错表格数据
sql_quality = 'select * from detail_supplier_quality'
YIELD_quality = pd.read_sql_query(sql=sql_quality, con=engine_detail_all)

#各供应商代码
list_code = (100105,100118,100120,100122,100140,100141,103195,103233,103275,103407,103695,104657,104661,104738,104656,105042,105078,105171)

#计算各供应商日合格率
YIELD = YIELD.fillna('-')
YIELD = YIELD[YIELD["过账日期"]==yesterday_format11+' 00:00:00']     #105筛出昨天的数据
YIELD_quality = YIELD_quality[YIELD_quality["退回日期"]==yesterday_format11]     #品质差错筛出昨天的数据

style_percent = easyxf(num_format_str='0.00%')   #定义为百分数
com_list,list_date,quality_list,rate_quality_list,standard_rate,leiji_number,leiji_quality_list,rate_leiji_quality_list = [[] for x in range(8)]    #定义空列表

#按供应商统计件数
for y in range(len(list_code)):
    each = YIELD[YIELD["供应商"]==list_code[y]]   #昨日各供应商订单量
    each_quality = YIELD_quality[YIELD_quality["供应商"]==list_code[y]]   #昨日各供应商品质差错件数
    each_leiji = leiji[leiji["供应商"]==list_code[y]]   #各供应商累计到昨日的总订单量
    each_leiji_quality = new[new["供应商"]==list_code[y]]  #各供应商累计到昨日的品质差错件数


    quality_list.append(each_quality.shape[0])  #统计昨日品质差错件数
    list_date.append(yesterday_format11)   #填入当天日期
    com_list.append(each.shape[0])   #统计供应商当日总件数
    leiji_number.append(each_leiji.shape[0])   #累计总订单数
    standard_rate.append(format(0.07/100,'.2%'))   #kpi
    leiji_quality_list.append(each_leiji_quality.shape[0])   #累计品质差错件数

#计算合格率
for j in range(len(list_code)):
    if com_list[j] == 0:
        rate_quality_list.append('-')
    else:
        rate_quality_list.append(format(quality_list[j]/com_list[j],'.2%'))
    if leiji_number[j] == 0:
        rate_leiji_quality_list.append('-')
    else:
        rate_leiji_quality_list.append(format(leiji_quality_list[j]/leiji_number[j],'.2%'))

def set_style(name, height, bold=False):
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = name
    font.bold = bold
    font.height = height
    style.font = font
    return style

#写入表格中
def financial_excel():
    f_dingding = xlwt.Workbook()
    sheet_w = f_dingding.add_sheet('每日合格率', cell_overwrite_ok=True)  #设置sheet名称
    row0 = ['日期','供应商', '当日订单数', '品质差错件数', '品差差错率','standard','总订单数','累计品差件数','累计品差率']   #设置第一行标题

    for n in range(0, len(row0)):
        sheet_w.write(0, n, row0[n], set_style('Time New Roman', 220, True))
    row_index = 1  # 第二行
    for m in range(len(list_code)):
        sheet_w.write(row_index, 0, str(list_date[m]), set_style('Time New Roman', 220))
        sheet_w.write(row_index, 1, str(list_code[m]), set_style('Time New Roman', 220))
        sheet_w.write(row_index, 2, str(com_list[m]), set_style('Time New Roman', 220))

        sheet_w.write(row_index, 3, str(quality_list[m]), set_style('Time New Roman', 220))
        sheet_w.write(row_index, 4, str(rate_quality_list[m]), style_percent)

        sheet_w.write(row_index, 5, str(standard_rate[m]), style_percent)
        sheet_w.write(row_index, 6, str(leiji_number[m]), set_style('Time New Roman', 220))
        sheet_w.write(row_index, 7, str(leiji_quality_list[m]), set_style('Time New Roman', 220))
        sheet_w.write(row_index, 8, str(rate_leiji_quality_list[m]), style_percent)
        row_index = row_index + 1

    f_dingding.save(file_path+yesterday_format+'各厂品质差错率.xls')
    print('写入成功！')

if __name__ == '__main__':
    financial_excel()

#将每日各工厂合格率追加到数据库的表格rate_each
YIELD = pd.read_excel(file_path+yesterday_format+'各厂品质差错率.xls')
print(YIELD)
YIELD.to_sql(name='rate_quality_each', con=engine_rate_all, if_exists='append', index=False, index_label=False)



#不合格原因计数----------------------------------------------------------------------------------------------------------------------
sql_supplier = 'select * from detail_supplier'
supplier = pd.read_sql_query(sql=sql_supplier, con=engine_detail_all)
supplier = supplier.drop_duplicates()
supplier.to_sql(name='detail_supplier',con=engine_detail_all, if_exists='replace', index=False, index_label=False)

recent = supplier[supplier['退回日期'].isin([yesterday_format11])]
list_supplier = list(set(recent['供应商'].tolist()))
df_combined = pd.DataFrame(columns=["供应商"])

for k in range(len(list_supplier)):
    recent_each = recent[recent["供应商"] == list_supplier[k]]
    display(recent_each)
    list_fen = recent_each['不合格原因'].tolist()

    total,baba,return_date = [[] for x in range(3)]
    for p in range(0, len(list_fen)):
        list_1 = list_fen[p].split('\\')
        total.extend(list_1)
    #     print(total)

    # 统计出现次数
    result = dict()
    for a in set(total):
        result[a] = total.count(a)

    df_result = pd.DataFrame([result])
    df_result = df_result.T
    df_result = df_result.reset_index()
    df_result.columns = ['不合格原因', '计数']

    for l in range(len(set(total))):
        baba.append(str(list_supplier[k]))
        return_date.append(yesterday_format)

    s = pd.Series(baba, name='供应商')
    v = pd.Series(return_date, name='退回日期')
    df_result = pd.concat([s, df_result, v], axis=1)

    df_combined = df_combined.append(df_result)
display(df_combined)
df_combined.to_sql(name='reason', con=engine_detail_all, if_exists='append', index=False, index_label=False)
