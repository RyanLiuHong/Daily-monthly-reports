import pyautogui
import time
import win32api
import win32con
import win32gui
import win32clipboard as w


def FindWindow(chatroom):
    win = win32gui.FindWindow('WeChatMainWndForPC', chatroom)
    print("找到窗口句柄：%x" % win)
    if win != 0:
        win32gui.ShowWindow(win, win32con.SW_SHOWMINIMIZED)
        win32gui.ShowWindow(win, win32con.SW_SHOWNORMAL)
        win32gui.ShowWindow(win, win32con.SW_SHOW)
        win32gui.SetWindowPos(win, win32con.HWND_TOP, 0, 0, 500, 700, win32con.SWP_SHOWWINDOW)
        win32gui.SetForegroundWindow(win)  # 获取控制
        time.sleep(1)
        tit = win32gui.GetWindowText(win)
        print('已启动【' + str(tit) + '】窗口')
    else:
        print('找不到【%s】窗口' % chatroom)
        exit()


# 设置和粘贴剪贴板
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

# 模拟发送动作
def SendMsg():
    win32api.keybd_event(18, 0, 0, 0)
    win32api.keybd_event(83, 0, 0, 0)
    win32api.keybd_event(18, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(83, 0, win32con.KEYEVENTF_KEYUP, 0)

# 模拟发送微信文本消息
def SendWxMsg(sendtext):
    #     # 先启动微信
    #     FindWindow('微信')
    #     time.sleep(1)
    #     # 定位到搜索框
    #     pyautogui.moveTo(143, 39)
    #     pyautogui.click()
    #     # 搜索微信
    #     ClipboardText(wxid)
    #     time.sleep(1)
    #     # 进入聊天窗口
    #     pyautogui.moveTo(155, 120)
    #     pyautogui.click()
    # 粘贴文本内容
    ClipboardText(sendtext)
    # 发送
    SendMsg()
    # 关闭微信窗口
    time.sleep(1)
    pyautogui.moveTo(683, 16)
    pyautogui.click()


# 模拟发送文件消息（图片、文档、压缩包等）
def SendWxFileMsg(wxid, imgpath):
    # 先启动微信
    FindWindow('微信')
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



list_code = (100105,100118,100120,100122,100140,100141,103195,103233,103275,103407,103695,103730,104657,104738,104656,105042,105078,105171,105630,105929)

# 文件所在路径
path = r'D:\1.office_syx\月报汇报' + '\\'

# 循环发送
for n in range(len(list_code)):
    # D:\1.office_syx\月报汇报\100120.xlsx # 发送文件消息（图片、文档、压缩包等）
    SendWxFileMsg(str(list_code[n]), path + str(list_code[n]) + '.xlsx')
    # 发送文本消息（微信号或微信昵称或备注，需要发送的文本消息）
    SendWxMsg('以上为贵司六月不良率及不良明细汇总，请查阅，如有疑异请于7月4日12点之前反馈我调整（@或私聊我)，7月4100120日后上报供应商管理部后则不可修改')
    print('{}已发送'.format(list_code[n]))
print('已发送完全部')