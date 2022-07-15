import pyautogui
import win32api
import win32con
import win32gui
import win32clipboard as w
import time

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


list_code = (141,120,105,407,233,105042,171,657)
list_name = ['华乐品质问题沟通','星宝缘品质问题沟通','凯沙琪品质问题沟通','金磨坊品质问题沟通','美恒诚品质问题沟通','启泰品质问题沟通','大新品质问题沟通','艺星品质问题沟通']

# 文件所在路径
path = r'D:' + '\\'

# 循环发送
for n in range(len(list_code)):
    # 发送文件消息（图片、文档、压缩包等）
    SendWxFileMsg(list_name[n], path + str(list_code[n]) + '.bmp')
    # 发送文本消息（微信号或微信昵称或备注，需要发送的文本消息）
    # SendWxMsg('贵司截至昨日不良率、错漏率，请查阅')   #截至昨日
    print('{}已发送'.format(list_code[n]))
print('已发送完全部')