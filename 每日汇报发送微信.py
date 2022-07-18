import pyautogui
import win32api
import win32con
import win32gui
import os
import platform
import subprocess
import win32clipboard as w
import datetime,time
from datetime import timedelta

yesterday = datetime.datetime.today()+timedelta(-1)
yesterday_format =yesterday.strftime('%Y.%m.%d')


def open_fp(fp: str):
    """
    打开文件或文件夹
    优点: 代码输入参数少, 复制粘贴即可使用, 支持在mac和win上使用, 打开速度快稳定;
    :param fp: 需要打开的文件或文件夹路径
    """
    systemType: str = platform.platform()  # 获取系统类型
    if 'mac' in systemType:  # 判断以下当前系统类型
        fp: str = fp.replace("\\", "/")  # mac系统下,遇到`\\`让路径打不开,不清楚为什么哈,觉得没必要的话自己可以删掉啦,18行那条也是
        subprocess.call(["open", fp])
    else:
        fp: str = fp.replace("/", "\\")  # win系统下,有时`/`让路径打不开
        os.startfile(fp)

def ClipboardText(ClipboardText):
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_UNICODETEXT, ClipboardText)
    w.CloseClipboard()
    time.sleep(1)
    win32api.keybd_event(17, 0, 0, 0)
    win32api.keybd_event(86, 0, 0, 0)
    win32api.keybd_event(86, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)

def FindWindow(WINDOW,chatroom):
    win = win32gui.FindWindow(WINDOW, chatroom)
    print("找到窗口句柄：%x" % win)
    if win != 0:
        win32gui.ShowWindow(win, win32con.SW_SHOWMINIMIZED)
        win32gui.ShowWindow(win, win32con.SW_SHOWNORMAL)
        win32gui.ShowWindow(win, win32con.SW_SHOW)
        if chatroom == 'Tableau - 每日汇报[已恢复]':
            win32gui.SetWindowPos(win, win32con.HWND_TOP, 0, 0, 1900, 1000, win32con.SWP_SHOWWINDOW)
        else:
            win32gui.SetWindowPos(win, win32con.HWND_TOP, 0, 0, 500, 700, win32con.SWP_SHOWWINDOW)
        win32gui.SetForegroundWindow(win)  # 获取控制
        time.sleep(1)
        tit = win32gui.GetWindowText(win)
        print('已启动【' + str(tit) + '】窗口')
    else:
        print('找不到【%s】窗口' % chatroom)
        if(WINDOW=='QT5QWindow' and chatroom=='文件恢复'):
            print("进入再一次文件恢复")
            time.sleep(1.5)
            FindWindow('Qt5QWindow','文件恢复')
        elif(WINDOW=='QT5QWindow' and chatroom=='登录以重新连接'):
            print("进入再一次登录")
            FindWindow('Qt5QWindow','登录以重新连接')
        else:exit()

# 模拟发送文件消息（图片、文档、压缩包等）
def SendWxFileMsg(wxid, imgpath):
    # 先启动微信
    FindWindow('WeChatMainWndForPC','微信')
    time.sleep(1)
    # 定位到搜索框
    pyautogui.moveTo(143, 39)
    pyautogui.click()
    # 搜索微信
    ClipboardText(wxid)
    time.sleep(1)
    # 进入聊天窗口
    pyautogui.moveTo(155, 120)
    pyautogui.click()
    # 选择文件
    pyautogui.moveTo(373, 570)
    pyautogui.click()
    ClipboardText(imgpath)
    time.sleep(1)
    pyautogui.moveTo(784, 509)
    pyautogui.click()
    # 发送
    SendMsg()
    # 关闭微信窗口
    # time.sleep(1)
    # pyautogui.moveTo(683, 16)
    # pyautogui.click()

def SendMsg():
    win32api.keybd_event(18, 0, 0, 0)
    win32api.keybd_event(83, 0, 0, 0)
    win32api.keybd_event(18, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(83, 0, win32con.KEYEVENTF_KEYUP, 0)

def SendWxMsg(sendtext):
    # 粘贴文本内容
    ClipboardText(sendtext)
    # 发送
    SendMsg()
    # 关闭微信窗口
    time.sleep(1)
    pyautogui.moveTo(683, 16)
    pyautogui.click()

if __name__ == '__main__':
    # 直接输入路径就可以使用, 绝对路径和相对路径都可以, 具体效果和鼠标双击文件或文件夹一样
    open_fp(fp='E:\dbeaver\dbeaver.exe')
    time.sleep(7)
    open_fp(fp='D:\\3.作图\每日汇报.twbx')
    time.sleep(5)
    FindWindow('QT5QWindow','文件恢复')
    time.sleep(1)

    # 定位到恢复
    pyautogui.moveTo(431, 669)
    pyautogui.click()
    time.sleep(6)

    # 定位到登陆账户
    FindWindow('Qt5QWindow','登录以重新连接')
    time.sleep(1)
    ClipboardText('DR003812')
    pyautogui.click()
    time.sleep(2)
    ClipboardText('DR003812')
    pyautogui.click()
    time.sleep(3)

    #定位到tableau，模板输入框
    FindWindow('Qt5QWindowIcon','Tableau - 每日汇报[已恢复]')
    pyautogui.moveTo(1488,129)
    pyautogui.click()
    list_code = (141, 120, 105, 407, 233, 105042, 171, 657)

    #导出tableau的图
    for n in range(len(list_code)):
        FindWindow('Qt5QWindowIcon', 'Tableau - 每日汇报[已恢复]')
        pyautogui.moveTo(1488, 129)
        pyautogui.click()
        ClipboardText(str(list_code[n]))
        time.sleep(1.5)
        # 点击对应模板
        pyautogui.moveTo(1506,171)
        time.sleep(1.5)
        pyautogui.click()
        time.sleep(1)
        # 定位到仪表盘
        pyautogui.moveTo(227,42)
        time.sleep(1)
        pyautogui.click()
        # 导出图像
        pyautogui.moveTo(263,242)
        time.sleep(1.5)
        pyautogui.click()
        time.sleep(1)
        ClipboardText(str(list_code[n]))
        pyautogui.moveTo(800,543)
        pyautogui.click()
        # 导出图像地址确定按钮
        pyautogui.moveTo(1000,530)
        time.sleep(1)
        pyautogui.click()

    # 文件所在路径
    path = r'D:' + '\\'
    list_name = ['华乐品质问题沟通', '星宝缘品质问题沟通', '凯沙琪品质问题沟通', '金磨坊品质问题沟通', '美恒诚品质问题沟通', '启泰品质问题沟通', '大新品质问题沟通', '艺星品质问题沟通']
    # 循环发送tableau图片

    # for n in range(len(list_code)):
    #     # 发送文件消息（图片、文档、压缩包等）
    #     SendWxFileMsg(list_name[n], path + str(list_code[n]) + '.bmp')
    #     # 发送文本消息（微信号或微信昵称或备注，需要发送的文本消息）
    #     # SendWxMsg('贵司截至昨日不良率、错漏率，请查阅')   #截至昨日
    #     print('{}已发送'.format(list_code[n]))
    # print('图片已发送完全部')

    path1 = r'D:\1.office_syx\不良\每日不良汇报' + '\\'
    list_code1 = (100105, 100118, 100120, 100122, 100140, 100141, 103195, 103233, 103407, 103695, 104657, 104738, 104656, 105042, 105078, 105171)
    # 循环发送当日excel报表

    # for n in range(len(list_code1)):
    #     # 发送文件消息（图片、文档、压缩包等）
    #     SendWxFileMsg(str(list_code1[n]), path1 + str(list_code1[n]) + '-' + yesterday_format + '不良.xlsx')
    #     # 发送文本消息（微信号或微信昵称或备注，需要发送的文本消息）
    #     SendWxMsg('贵司截至昨日不良率、错漏率，请查阅')  # 截至昨日
    #     print('{}已发送'.format(list_code1[n]))

    print('微信每日汇报已完成')
    time.sleep(2)

    FindWindow('Qt5QWindowIcon','Tableau - 每日汇报[已恢复]')
    # 推出每日汇报[已恢复]
    pyautogui.moveTo(1869,6)
    pyautogui.click()
    pyautogui.moveTo(871,535)
    pyautogui.click()
    pyautogui.moveTo(795,539)
    pyautogui.click()
    pyautogui.moveTo(1005, 530)
    pyautogui.click()
    # 退出每日汇报
    FindWindow('Qt5QWindowIcon', 'Tableau - 每日汇报')
    pyautogui.moveTo(483,8)
    pyautogui.click()
    FindWindow('SWT_Window0','DBeaver 21.3.2 - detail_105')
    pyautogui.moveTo(483,8)
    pyautogui.click()



