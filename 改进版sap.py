import sys, win32com.client,os
import win32api, win32gui, win32con,subprocess
import datetime,time
from datetime import timedelta

today = time.strftime("%Y-%m-%d")  #2022-04-19
year = datetime.datetime.now().year
yesterday = datetime.datetime.today()+timedelta(-1)
yesterday_format = yesterday.strftime('%Y.%m.%d')
month = yesterday.month
yccc = '2022.06.01'


def Main():

#1.登入sap-----------------------------------------------------------------------------
    print('准备进入sap')
    try:
        sap_app = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"  # 您的saplogon程序本地完整路径
        subprocess.Popen(sap_app)
        print('1')

        time.sleep(2)

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

        # # 出现版权点击确定，没有版权提示直接进行下一步
        # try:
        #     x = session.findById("wnd[1]").text
        #     # print(x)
        #     if 'Copyright' in x:
        #         session.findById("wnd[1]/tbar[0]/btn[0]").press()
        # except:
        #     pass

        # # 出现二级密码登录，不出现跳过
        # try:
        #     session.findById("wnd[1]/usr/txtGS_OUT-ID").text = username2
        #     session.findById("wnd[1]/usr/pwdGS_OUT-PW").text = password2
        #     session.findById("wnd[1]/usr/pwdGS_OUT-PW").setFocus()
        #     session.findById("wnd[1]/usr/pwdGS_OUT-PW").caretPosition = 10
        #     session.findById("wnd[1]/usr/btnLOGIN").press()
        # except:
        #     print('no second user')

        #检验是否有文件夹如果没有需要创建
        os.makedirs("D:\\1.office_syx\\不良\\"+str(year)+"."+str(month)+"不良",exist_ok=True)
        #文件保存路径
        path = "D:\\1.office_syx\\不良\\"+str(year)+"."+str(month)+"不良\\"

    #2.打印出105-----------------------------------------------------------------------------
        # 转到zmm105
        time.sleep(10)
        session.findById("wnd[0]").resizeWorkingPane(151, 27, 0)
        session.findById("wnd[0]/tbar[0]/okcd").text = "zmm105"
        session.findById("wnd[0]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/usr/ctxtP_ZVERID").text = "B1"
        session.findById("wnd[0]/usr/txtS_MONAT-LOW").text = month
        session.findById("wnd[0]/usr/txtS_MONAT-HIGH").text = month
        session.findById("wnd[0]/usr/ctxtS_BUDAT-LOW").text =yccc
        session.findById("wnd[0]/usr/ctxtS_BUDAT-HIGH").text = yesterday_format
        session.findById("wnd[0]/usr/txtS_MONAT-HIGH").setFocus()
        session.findById("wnd[0]/usr/txtS_MONAT-HIGH").caretPosition = 2
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        print('已进入105')
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&XXL")
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = path # 导出表格
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = yesterday_format + "不良明细.XLSX"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 4
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        sapInfo = session.findById("wnd[0]/sbar").text # 由于SAP脚本自身的特性，当程序读到左下角消息时，意味着数据已经传输完成！
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/btn[15]").press()  # 退出到上一级
        session.findById("wnd[0]/tbar[0]/btn[15]").press()

        excel1 = win32gui.FindWindow("XLMAIN", yesterday_format+'不良明细 - Excel')  # 关闭窗口
        win32gui.SendMessage(excel1, win32con.WM_CLOSE, None, None)
        print('已关闭105')

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
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = yesterday_format + "不良率.XLSX"
        session.findById("wnd[1]/tbar[0]/btn[11]").press()  # 点击【替换】
        print('已导出084')
        sapInfo = session.findById("wnd[0]/sbar").text  # 由于SAP脚本自身的特性，当程序读到左下角消息时，意味着数据已经传输完成！
        session.findById("wnd[0]").maximize()
        time.sleep(10)
        session.findById("wnd[0]/tbar[0]/btn[15]").press()  # 退出到上一级
        session.findById("wnd[0]/tbar[0]/btn[15]").press()

        excel3 = win32gui.FindWindow("XLMAIN", yesterday_format+'不良率 - Excel')  # 关闭窗口
        win32gui.SendMessage(excel3, win32con.WM_CLOSE, None, None)
        print('已关闭084')

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


