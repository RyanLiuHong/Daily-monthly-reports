#导入包-------------------------------------------------------------
import pandas as pd
from IPython.display import display
import warnings
import datetime,time
from datetime import timedelta
from sqlalchemy import create_engine
from xlwt import easyxf,Workbook
import xlwt,xlrd,os
warnings.filterwarnings("ignore")

#1、时间定义-------------------------------------------------------------
today = time.strftime("%Y-%m-%d")  #2022-04-19
yesterday = datetime.datetime.today()+timedelta(-1)
yesterday_format = '2022.06.12'#yesterday.strftime('%Y.%m.%d')
month = yesterday.month
md = '6.1-'+str(month)+'.'+str(yesterday.day)
yesterday_format11 ='2022-06-12'  #yesterday.strftime('%Y-%m-%d')

#数据库
engine_detail_all= create_engine('mysql+pymysql://root:DR003812@localhost:3306/detail_all?charset=utf8')
engine_rate_all= create_engine('mysql+pymysql://root:DR003812@localhost:3306/rate_all?charset=utf8')

# 文件保存路径
file_path = "D:\\1.office_syx\\不良\\2022." + str(month) + "不良\\"

#2、数据导入-------------------------------------------------------------
oneofive = pd.read_excel(file_path+yesterday_format+'-105.xlsx', index_col=0)
oneofive.to_sql(name='detail_105',con=engine_detail_all,if_exists='append',index=False,index_label=False)

# xlsx转为csv
def xlsx_to_csv_pd(path):
    data_xls = pd.read_excel(path+'.xlsx', index_col=0)
    data_xls.to_csv(path+'.csv', encoding='utf_8_sig')

if __name__ == '__main__':
    xlsx_to_csv_pd(file_path + yesterday_format + '不良明细')
    xlsx_to_csv_pd(file_path + yesterday_format + '不良率')

# 数据导入
bs = pd.read_csv(file_path + yesterday_format + '不良明细.csv', encoding='utf-8')
br = pd.read_csv(file_path + yesterday_format + '不良率.csv', encoding='utf-8', thousands=',')
leiji = pd.read_excel(file_path+'累计-105.xlsx', index_col=0)

# 不良明细汇报数据处理
bs = bs[bs['退回日期'].isin([yesterday_format11])]  #筛选包含昨天日期的行   yesterday_format11

#将105、不良率及不良明细导入到sql
bs.to_sql(name='detail_com',con=engine_detail_all,if_exists='append',index=False,index_label=False)
br.to_sql(name='rate_all', con=engine_rate_all, if_exists='append', index=False, index_label=False)

#3、数据处理-------------------------------------------------------------

#不良明细分表处理
bs_gys = bs[bs['责任归属'].isin(['供应商责任'])]   #筛出供应商责任
bs_gys = bs_gys.reset_index(drop = True)
bs_ws = bs[bs['责任归属'].isin(['我司责任'])]    #筛出我司原因

# 筛选出品质差错
ls = ['漏刻字', '刻字错', '石重低配', '镶错石', '款式错', '材质错', '手寸错', '成色不足']
new= pd.DataFrame(columns = ['采购订单','定制编码','供应商','款号','款式名称','批次','质检结果','不合格原因','责任归属','质检地点','退回日期','订单类型描述','质检人'])

for j in range(0,bs_gys.shape[0]):
    for i in ls:
        if i in str(bs_gys.iloc[j,7]):
            TTT = pd.DataFrame(bs_gys.loc[j]).T
            display(TTT)
            new = new.append(TTT,ignore_index=True)

ls1 = ['/漏刻字', '/刻字错', '/石重高配','/单反石','/镶错石','/款式错','/调乱货','/材质错','/链不符','/手寸错','/版型不符','/成色不足','/货重不符','/来货信息不符','/款式错','/材质错','/链不符','/货重不符']
new1 = pd.DataFrame(columns = ['采购订单','定制编码','供应商','款号','款式名称','批次','质检结果','不合格原因','责任归属','质检地点','退回日期','订单类型描述','质检人'])

for j in range(0,bs_gys.shape[0]):
    for i in ls1:
        if i in str(bs_gys.iloc[j,7]):
            TTT = pd.DataFrame(bs_gys.loc[j]).T
            display(TTT)
            new1= new1.append(TTT,ignore_index=True)

#导入到数据库
bs_gys.to_sql(name='detail_supplier',con=engine_detail_all, if_exists='append', index=False, index_label=False)    #供应商责任
new.to_sql(name='detail_supplier_quality', con=engine_detail_all, if_exists='append', index=False, index_label=False)    #供应商责任-品质差错
bs_ws.to_sql(name='detail_our', con=engine_detail_all, if_exists='append', index=False, index_label=False)   #我司责任


#不良率数据处理
br1=br.drop(columns=["标准总不良率","实际总不良率环比","标准工艺不良率","工艺标准达成率","工艺不良率环比","标准错漏率","错漏标准达成率","错漏超出件数","实际错漏率环比"])    #删除不需要的列
br1['总订单量'] = pd.to_numeric(br1['总订单量'],errors='coerce')
br1['实际总不良件数'] = pd.to_numeric(br1['实际总不良件数'],errors='coerce')
br1['实际工艺不良件数'] = pd.to_numeric(br1['实际工艺不良件数'],errors='coerce')
br1['实际错漏件数'] = pd.to_numeric(br1['实际错漏件数'],errors='coerce')
br1.head()

#累计105数据处理
leiji_gys = leiji[leiji['责任归属'].isin(['供应商责任'])]

#供应商列表
list_code = (100105,100118,100120,100122,100140,100141,103195,103233,103275,103407,103695,104657,104661,104738,104656,105042,105078,105171)

style_percent = easyxf(num_format_str='0.00%')   #定义为百分数
fine_list,dateddd_list,standar_rate,com_list,day_list,rate_day_list = [[] for x in range(6)]   #定义空列表

YIELD = oneofive[oneofive["过账日期"]==yesterday_format11]   #筛选出105表格中过账日期为昨天的数据

#设置字体格式
def set_style(name, height, bold=False):
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = name
    font.bold = bold
    font.height = height
    style.font = font
    return style

#4、输出-------------------------------------------------------------
#不良率计算

#筛选出各供应商数据
for p in range(0,len(list_code)):
    bbb = br1[br1["供应商"] == list_code[p]]
    each = YIELD[YIELD["供应商"] == list_code[p]]
    ccc = bs_gys[bs_gys["供应商"] == list_code[p]]
    ddd = leiji[leiji["供应商"] == list_code[p]]
    ddd_gys = ddd[ddd['责任归属'].isin(['供应商责任'])]
    eee = new1[new1["供应商"] == list_code[p]]

    try:
        bbb.loc['sum'] = bbb.iloc[0:500, [4, 5, 7, 9]].sum(axis=0)   #总订单量、实际总不良件数、实际工艺不良件数、实际错漏件数求和

        change = ('总订单量', '实际总不良件数', '实际工艺不良件数', '实际错漏件数')
        for j in range(0, 4):
            bbb[change[j]] = bbb[change[j]].astype("Int64")

        bbb.iloc[bbb.shape[0] - 1, 6] = format(bbb.iloc[bbb.shape[0] - 1, 5] / bbb.iloc[bbb.shape[0] - 1, 4], '.2%')   #计算实际总不良率
        bbb.iloc[bbb.shape[0] - 1, 8] = format(bbb.iloc[bbb.shape[0] - 1, 7] / bbb.iloc[bbb.shape[0] - 1, 4], '.2%')   #计算实际工艺不良率
        bbb.iloc[bbb.shape[0] - 1, 10] = format(bbb.iloc[bbb.shape[0] - 1, 9] / bbb.iloc[bbb.shape[0] - 1, 4], '.2%')   #计算实际错漏率
        bbb.iat[bbb.shape[0] - 1, 2] = md
        bbb.iat[bbb.shape[0] - 1, 3] = '汇总'
        bbb = bbb.fillna('-')   #空值用-填充
        display(bbb)

        fine_list.append(format(ddd_gys.shape[0]/ddd.shape[0],'.2%'))  #累计不合格率
        dateddd_list.append(yesterday_format11)   #日期
        standar_rate.append(format(99.1/100,'.2%'))   #标准合格率
        com_list.append(each.shape[0])  # 统计供应商当日总件数
        day_list.append(ccc.shape[0])   #当日不合格件数


        with pd.ExcelWriter(r"D:/1.office_syx/不良/每日不良汇报/{}不良.xlsx".format(list_code[p])) as writer:   #写出到每日汇报的表格
            bbb.to_excel(writer, sheet_name='不良率', index=False)
            ccc.to_excel(writer, sheet_name='不良明细',index = False)
            eee.to_excel(writer, sheet_name='错漏明细', index=False)

    except:
        print("遇到错误啦！")


for t in range(len(list_code)):
    if com_list[t] == 0:
        rate_day_list.append('-')    #如果当日无订单，用-填充
    else:
        rate_day_list.append(format((com_list[t]-day_list[t])/com_list[t],'.2%'))    #如果有订单，正常计算

#写入表格中
def financial_excel():
    f_dingding = xlwt.Workbook()
    sheet_w = f_dingding.add_sheet('累计今日合格率', cell_overwrite_ok=True)  #设置sheet名称
    row0 = ['日期','供应商', '当日订单数','当日不合格件数','当日合格率', '累计不合格率','标准合格率']   #设置第一行标题

    for n in range(0, len(row0)):
        sheet_w.write(0, n, row0[n], set_style('Time New Roman', 220, True))   #填入标题，即第一行
    row_index = 1  # 第二行
    for m in range(len(list_code)):
        sheet_w.write(row_index, 0, str(dateddd_list[m]), set_style('Time New Roman', 220))   #日期
        sheet_w.write(row_index, 1, str(list_code[m]), set_style('Time New Roman', 220))    #供应商列表
        sheet_w.write(row_index, 2, str(com_list[m]), set_style('Time New Roman', 220))   #当日订单数
        sheet_w.write(row_index, 3, str(day_list[m]), set_style('Time New Roman', 220))   #当日不合格件数
        sheet_w.write(row_index, 4, str(rate_day_list[m]), style_percent)   #当日合格率
        sheet_w.write(row_index, 5, str(fine_list[m]), style_percent)    #累计不合格率
        sheet_w.write(row_index, 6, str(standar_rate[m]), style_percent)   #标准合格率

        row_index = row_index + 1

    f_dingding.save(file_path+'各厂累计'+yesterday_format+'合格率.xls')
    accumulate = pd.read_excel(file_path+'各厂累计'+yesterday_format+'合格率.xls')
    accumulate.to_sql(name='rate_each', con=engine_rate_all, if_exists='append', index=False, index_label=False)   #写入数据库
    print('写入成功！')

financial_excel()


#6、删除不需要的文件
os.remove(file_path + yesterday_format + '不良明细.xlsx')
os.remove(file_path + yesterday_format + '不良率.xlsx')