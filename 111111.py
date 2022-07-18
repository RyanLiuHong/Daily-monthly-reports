import os
from time import sleep
def run():
 while 1:
    # os.system("python ./改进版sap.py")
    # print("等一会会儿吧...")
    # sleep(10)
    #
    # os.system("python ./改进版tosql.py")
    # print("再等一会会儿吧 ")
    # sleep(8)
    #
    # os.system("python ./改进版rate.py")
    # print("再再再等一会会儿吧 ")
    # sleep(6)

    # print('现在要开始发微信啦，千万不要动到鼠标，会出大事！！！')
    # os.system("python ./4_wechat11.py")
    # sleep(5)

    # print('现在要开始发tableau每日汇报，不要动鼠标喔')
    # # os.system("python ./每日汇报发送微信.py")
    # sleep(3)
    #
    print('现在开始更新共享盘文件啦！！！')
    os.system("python ./日常更新共享盘文件.py")
    sleep(5)
    exit()

run()