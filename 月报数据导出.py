import sys, win32com.client
import win32api, win32gui, win32con, win32ui, time, os, subprocess
import datetime,time
import pandas as pd
from datetime import timedelta

print('怎么回事')
today = time.strftime("%Y-%m-%d")  #2022-04-19
year = 2022
month = 6
ym = '2022.06'
yc = '2022.06.01'
yw = '2022.06.30'

# 文件保存路径
path = "D:\\1.office_syx\\不良\\2022.06不良月报\\"

def Main():

#1.登入sap-----------------------------------------------------------------------------
    print('准备进入sap')
    try:
        sap_app = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"  # 您的saplogon程序本地完整路径
        subprocess.Popen(sap_app)

        time.sleep(1)

        flt = 0
        while flt == 0:
            try:
                hwnd = win32gui.FindWindow(None, "SAP Logon 770")
                flt = win32gui.FindWindowEx(hwnd, None, "Edit", None)
                time.sleep(3)

            except:
                print('except============>')
                time.sleep(0.5)
        win32gui.SendMessage(flt, win32con.WM_SETTEXT, None, "800")
        win32gui.SendMessage(flt, win32con.WM_KEYDOWN, win32con.VK_RIGHT, 0)
        win32gui.SendMessage(flt, win32con.WM_KEYUP, win32con.VK_RIGHT, 0)
        time.sleep(2)
        # 登录GUI界面

        time.sleep(3)
        dlg = win32gui.FindWindowEx(hwnd, None, "Button", None)
        win32gui.SendMessage(dlg, win32con.WM_LBUTTONDOWN, 0)
        win32gui.SendMessage(dlg, win32con.WM_LBUTTONUP, 0)

        time.sleep(2)
        SapGuiAuto = win32com.client.GetObject("SAPGUI")

        print(SapGuiAuto)
        print(type(SapGuiAuto))

        if not type(SapGuiAuto) == win32com.client.CDispatch:
            return
        application = SapGuiAuto.GetScriptingEngine

        while not application.Children.Count:
            time.sleep(5)
            win32gui.SendMessage(dlg, win32con.WM_LBUTTONDOWN, 0)
            win32gui.SendMessage(dlg, win32con.WM_LBUTTONUP, 0)

            SapGuiAuto = win32com.client.GetObject("SAPGUI")
            application = SapGuiAuto.GetScriptingEngine

        if not type(application) == win32com.client.CDispatch:
            SapGuiAuto = None
            return
        connection = application.Children(0)
        if not type(connection) == win32com.client.CDispatch:
            application = None
            SapGuiAuto = None
            return
        time.sleep(5)
        flag = 0
        while flag == 0:
            try:
                session = connection.Children(0)
                flag = 1
            except:
                time.sleep(5)
        # print('type session', type(session))
        if not type(session) == win32com.client.CDispatch:
            connection = None
            application = None
            SapGuiAuto = None
            return

        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "QM01"  # SAP登陆用户名
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "Dr147852369"  # SAP登陆密码
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").setFocus()
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 12
        session.findById("wnd[0]").sendVKey(0)

        # 出现多用户登录
        # License Information for Multiple Logon

        multi_logon_text = session.findById("wnd[1]").text

        if session.children.count > 1:
            session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select()
            session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").setFocus()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()

        print('已登入系统')

    #2.打印出105-----------------------------------------------------------------------------
        # 转到zmm105
        time.sleep(10)
        session.findById("wnd[0]").resizeWorkingPane(151, 27, 0)
        session.findById("wnd[0]/tbar[0]/okcd").text = "zmm105"
        session.findById("wnd[0]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/usr/ctxtP_ZVERID").text = "B1"
        session.findById("wnd[0]/usr/txtS_MONAT-LOW").text = month
        session.findById("wnd[0]/usr/txtS_MONAT-HIGH").text = month
        session.findById("wnd[0]/usr/ctxtS_BUDAT-LOW").text =yc
        session.findById("wnd[0]/usr/ctxtS_BUDAT-HIGH").text = yw
        session.findById("wnd[0]/usr/txtS_MONAT-HIGH").setFocus()
        session.findById("wnd[0]/usr/txtS_MONAT-HIGH").caretPosition = 2
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        print('已进入105')
        time.sleep(10)
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&XXL")
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(10)
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = path
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = ym+"-105.XLSX"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 4
        time.sleep(10)
        session.findById("wnd[1]/tbar[0]/btn[11]").press()

        print('已导出105')
        sapInfo = session.findById("wnd[0]/sbar").text
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/btn[15]").press()  # 退出到上一级
        time.sleep(10)

        #导出105表
        data_xls = pd.read_excel(path + ym+"-105.XLSX")
        time.sleep(10)
        data_xls = data_xls.fillna('-')
        data_xls = data_xls.loc[~data_xls['不良原因'].isin(['-'])]  # 筛选出不良原因中有值的行
        print(data_xls.shape)
        list = data_xls['批次'].tolist()
        print(len(list))

        # 提出批次放入txt中
        with open(path+"picimonth.txt", 'w') as f1:
            # 方法一：
            for line in list:
                f1.write(line + '\n')

        excel1 = win32gui.FindWindow("XLMAIN", ym+'-105 - Excel')  # 关闭窗口
        win32gui.SendMessage(excel1, win32con.WM_CLOSE, None, None)

    #打印不良原因前三项
        session.findById("wnd[0]/usr/radRB_03").setFocus()
        session.findById("wnd[0]/usr/radRB_03").select()
        session.findById("wnd[0]/usr/txtS_MONAT-LOW").text = 5
        session.findById("wnd[0]/usr/txtS_MONAT-HIGH").text = 5          #ggggdhghtyd
        session.findById("wnd[0]/usr/ctxtS_BUDAT-LOW").text =''
        session.findById("wnd[0]/usr/ctxtS_BUDAT-HIGH").text = ''
        session.findById("wnd[0]/usr/txtS_MONAT-HIGH").setFocus()
        session.findById("wnd[0]/usr/txtS_MONAT-HIGH").caretPosition = 2
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&XXL")
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = path
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = ym+"不良原因前三项.XLSX"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 28
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()

        excel8 = win32gui.FindWindow("XLMAIN", ym+'不良原因前三项 - Excel')  # 关闭窗口
        win32gui.SendMessage(excel8, win32con.WM_CLOSE, None, None)
        print('已关闭105')

    #3.打印不良明细-----------------------------------------------------------------------------
        # 转到zmm045B
        time.sleep(10)
        session.findById("wnd[0]").resizeWorkingPane(151, 27, 0)
        session.findById("wnd[0]/tbar[0]/okcd").text = "zmm045B"
        session.findById("wnd[0]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/usr/btn%_S_CHARG_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[23]").press()
        session.findById("wnd[2]/usr/ctxtDY_PATH").text = path
        session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = "picimonth.txt"
        session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 8
        session.findById("wnd[2]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        time.sleep(10)
        print('已进入045')
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&XXL")
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = path
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = ym+"不良明细.XLSX"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 12
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        print('已导出045')
        time.sleep(10)
        sapInfo = session.findById("wnd[0]/sbar").text
        session.findById("wnd[0]").maximize()
        time.sleep(10)
        session.findById("wnd[0]/tbar[0]/btn[15]").press()  # 退出到上一级
        session.findById("wnd[0]/tbar[0]/btn[15]").press()

        excel2 = win32gui.FindWindow("XLMAIN", ym+'不良明细 - Excel')  # 关闭窗口
        win32gui.SendMessage(excel2, win32con.WM_CLOSE, None, None)
        print('已关闭045')

    #4.打印不良率-----------------------------------------------------------------------------
        # 转到zmm084B表
        session.findById("wnd[0]").resizeWorkingPane(151, 27, 0)
        session.findById("wnd[0]/tbar[0]/okcd").text = "ZMM084B"
        session.findById("wnd[0]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/usr/radRB_09").select()
        session.findById("wnd[0]/usr/txtP_MJAHR").text = year
        session.findById("wnd[0]/usr/txtS_MONAT-LOW").text = month
        session.findById("wnd[0]/usr/txtS_MONAT-HIGH").text = month
        session.findById("wnd[0]/usr/txtS_MONAT-HIGH").setFocus()
        session.findById("wnd[0]/usr/txtS_MONAT-HIGH").caretPosition = 2
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        print('已进入084')
        time.sleep(10)
        
        # 导出不良率
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&XXL")
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = path # 保存路径
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = ym+ "不良率.XLSX"
        session.findById("wnd[1]/tbar[0]/btn[11]").press()  # 点击【替换】
        print('已导出084')
        sapInfo = session.findById("wnd[0]/sbar").text  # 由于SAP脚本自身的特性，当程序读到左下角消息时，意味着数据已经传输完成！
        session.findById("wnd[0]").maximize()
        time.sleep(10)
        session.findById("wnd[0]/tbar[0]/btn[15]").press()  # 退出到上一级
        session.findById("wnd[0]/tbar[0]/btn[15]").press()

        excel3 = win32gui.FindWindow("XLMAIN", ym+'不良率 - Excel')  # 关闭窗口
        win32gui.SendMessage(excel3, win32con.WM_CLOSE, None, None)
        print('已关闭084')

    # 5.打印降级入库-----------------------------------------------------------------------------
        # 转到zmm061表
        session.findById("wnd[0]").resizeWorkingPane(151, 27, 0)
        session.findById("wnd[0]/tbar[0]/okcd").text = "zmm061"
        session.findById("wnd[0]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/usr/ctxtS_ZJIEG-LOW").text = "3"
        session.findById("wnd[0]/usr/ctxtS_ZEREN-LOW").text = "2"
        session.findById("wnd[0]/usr/ctxtS_STATU-LOW").text = "1"
        session.findById("wnd[0]/usr/ctxtS_ZCZRQ-LOW").text = "2022.06.01"
        session.findById("wnd[0]/usr/ctxtS_ZCZRQ-HIGH").text = "2022.06.30"
        session.findById("wnd[0]/usr/ctxtS_ZCZRQ-HIGH").setFocus()
        session.findById("wnd[0]/usr/ctxtS_ZCZRQ-HIGH").caretPosition = 10
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = "MATNR"
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&XXL")
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = path
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = ym+"降级入库.XLSX"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()

        excel4 = win32gui.FindWindow("XLMAIN", ym+'降级入库 - Excel')  # 关闭窗口
        win32gui.SendMessage(excel4, win32con.WM_CLOSE, None, None)

    #6.打印退货明细-----------------------------------------------------------------------------------------
        session.findById("wnd[0]").resizeWorkingPane(151, 27, 0)
        session.findById("wnd[0]/tbar[0]/okcd").text = "mb51"
        session.findById("wnd[0]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = "DC01"
        session.findById("wnd[0]/usr/ctxtBWART-LOW").text = "122"
        session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = "2022.06.01"
        session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = "2022.06.30"
        session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").setFocus()
        session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").caretPosition = 10
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&XXL")
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = path
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = ym+"退货明细.XLSX"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()

        excel5 = win32gui.FindWindow("XLMAIN", ym+'退货明细 - Excel')  # 关闭窗口
        win32gui.SendMessage(excel5, win32con.WM_CLOSE, None, None)

        #关闭sap
        session.findById("wnd[0]").resizeWorkingPane(151, 27, 0)
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
        session.findById("wnd[0]/tbar[0]/btn[0]").press()


    except Exception as e:
        print(e)
        print(sys.exc_info()[0])

    finally:
        hwnd11 = win32gui.FindWindow(None, "SAP Logon 770")
        win32gui.SendMessage(hwnd11, win32con.WM_CLOSE, None, None)

if __name__ == "__main__":
    Main()

os.remove(path + ym + '-105.xlsx')


