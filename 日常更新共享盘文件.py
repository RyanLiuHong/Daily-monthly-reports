import datetime
from datetime import timedelta
import win32com.client,time
import shutil

yesterday = datetime.datetime.today()+timedelta(-1)
yesterday_format = yesterday.strftime('%Y.%m.%d')
year = yesterday.year
month = yesterday.month

# 将每日不良率更新
xls = win32com.client.Dispatch("Excel.Application")
wb=xls.Workbooks.Open('\\'+'\\'+'192.168.20.249\Shares\\4、质检部\\5_品质提升组\\每日不良率.xlsm')
xls.Application.Run("自动更新数据")  #执行excel中的VBA
time.sleep(1)
wb.Save()
wb.Close()
xls.Application.Quit()
# 将昨日不良率及不良明细添加进共享文件夹中
old_path1 = 'D:\\1.office_syx\不良\\'+str(year)+'.'+str(month)+'不良\\'+yesterday_format+'不良率.csv'
old_path2 = 'D:\\1.office_syx\不良\\'+str(year)+'.'+str(month)+'不良\\'+yesterday_format+'不良明细.xlsx'
path_list=[old_path1,old_path2]
print(old_path1,old_path2)
new_path = '\\'+'\\'+'192.168.20.249\Shares\\4、质检部\\5_品质提升组\\'+str(year)+'年'+str(month)+'月'
for i in path_list:
    shutil.copy(i, new_path)

