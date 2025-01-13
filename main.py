"""
作者: Sugobet
QQ: 321355478
Github:
脚本开源免费，使用前请先查阅readme.md文档
有任何问题请QQ联系，有必要时将会收取一定费用
二次开发
狼元星
QQ:3186444719
Github:
https://github.com/yuanxinglang
bilibili:
https://space.bilibili.com/516549721
"""

import os
import threading
import time
from ctypes import *
import cv2
import uiautomation as auto
import win32clipboard
import win32con
from PIL import Image, ImageFile, UnidentifiedImageError
from uiautomation.uiautomation import Bitmap

ImageFile.LOAD_TRUNCATED_IMAGES = True

wx_group_name = '狼元星'  # 微信群名， 留空则不发微信     请务必将要发送的群或人的这个聊天窗口设置成独立窗口中打开，并且不要最小化
wx_context = '星系警告!!!'  # 要发送的微信消息
con_val = 0.5  # 阈值：程序执行一遍的间隔时间，单位：秒
# 路径请勿出现中文
path = 'E:/GitHub/EVE_A_Eye_NEW'  # 脚本目录绝对路径 请将复制过来的路径的反斜杠修改成斜杠！
devices = {  # 模拟器地址，要开几个预警机就填对应预警机的模拟器的地址， 照抄
    'S1': [  # 星系名        请勿出现中文或中文字符以及特殊字符
        '127.0.0.1:5559',  # cmd输入adb devices查看模拟器地址
        False  # 没卵用，但不能少也不能改
    ],
}
game_send_position = {  # 从聊天框中第二个频道开始数，即系统频道之后为第二频道
    '第二频道': '38 117',  # 本地频道
    '第三频道': '38 170',  # 军团频道
    '第四频道': '38 223',
    '第五频道': '38 278',
    '第六频道': '38 332',
    '第七频道': '38 382'
}
sendTo = game_send_position['第三频道']  # 默认发送军团频道

mutex = threading.Lock()


def set_clipboard_file(paths):
    try:
        im = Image.open(paths)
        im.save('1.bmp')
        a_string = windll.user32.LoadImageW(0, r"1.bmp", win32con.IMAGE_BITMAP, 0, 0, win32con.LR_LOADFROMFILE)
    except UnidentifiedImageError:
        set_clipboard_file(paths)
        return

    if a_string != 0:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_BITMAP, a_string)
        win32clipboard.CloseClipboard()
        return
    print('图片载入失败')


def send_msg(content, msg_type=1):
    wechat_window = auto.WindowControl(
        searchDepth=1, Name=f"{wx_group_name}")
    wechat_window.SetActive()
    edit = wechat_window.EditControl()
    if msg_type == 1:
        auto.SetClipboardText(content)
    elif msg_type == 2:
        auto.SetClipboardBitmap(Bitmap.FromFile(content))
    elif msg_type == 3:
        set_clipboard_file(content)
    edit.SendKeys('{Enter}')
    edit.SendKeys('{Ctrl}v')
    edit.SendKeys("{Enter}")


def start():
    # 启动时重置  ‘’图片
    with open(f'{path}/tem/list.png', 'rb') as sc1:
        con = sc1.read()
        for k in devices:
            f = open(f'{path}/new_{k}_list.png', 'wb')
            f.write(con)
            f.close()

    with open(f'{path}/tem/playerList.png', 'rb') as sc:
        con = sc.read()
        for k in devices:
            f = open(f'{path}/old_{k}_playerList.png', 'wb')
            f.write(con)
            f.close()
            f = open(f'{path}/new_{k}_playerList.png', 'wb')
            f.write(con)
            f.close()

    # 监听线程
    for k in devices:
        t = threading.Thread(target=listening, args=(k,))
        t.start()

    print('Started')
    context = f"预警系统已上线，监测星系列表：\n {devices.keys()}"
    mutex.acquire()
    send_msg(context, msg_type=1)
    mutex.release()


def screenc(filename, num):
    os.system(f'adb -s {devices[filename][0]} exec-out screencap -p > {filename}_{num}.png')


def crop(x1, y1, x2, y2, sc_file_name, sv_file_name):
    try:
        img = Image.open(sc_file_name)
        re = img.crop((x1, y1, x2, y2))
        re.save(sv_file_name)
        img.close()
        re.close()
    except IOError:
        print("无法打开或读取文件")
    except ValueError:
        print("裁剪区域无效")
    except Exception as e:
        print("其他错误"+str(e))
        return


def load_image(img1, img2):
    i1 = cv2.imread(img1, 0)
    i2 = cv2.imread(img2, 0)
    return i1, i2


def if_img_i(src, mp):
    try:
        res = cv2.matchTemplate(src, mp, cv2.TM_CCOEFF_NORMED)
    except cv2.error as e:  # 捕获 OpenCV 相关的错误
        print("OpenCV 错误: "+str(e))
        return False, 0.999
    except ValueError as e:  # 捕获值错误，例如传入的参数类型不正确
        print("值错误: "+str(e))
        return False, 0.999
    except Exception as e:  # 捕获其他未预料到的异常
        print("其他错误: "+str(e))
        return False, 0.999

    _, mac_v, _, _ = cv2.minMaxLoc(res)
    if mac_v < 0.99:
        return True, mac_v
    return False, mac_v


def send_game_massage(tag):
    str1 = f'adb -s {devices[tag][0]} '
    os.system(str1 + 'shell input tap 211 478')  # 进入聊天频道
    time.sleep(1)
    os.system(str1 + f'shell input tap {sendTo}')  # 进入目标频道
    time.sleep(0.5)
    os.system(str1 + 'shell input tap 266 520')  # 进入更多
    time.sleep(0.2)
    os.system(str1 + 'shell input tap 870 515')  # 进入汇报
    time.sleep(0.2)
    os.system(str1 + 'shell input tap 75 290')  # 进入情报
    time.sleep(0.2)
    os.system(str1 + 'shell input tap 235 435')  # 写入侦查(汇报人数)
    time.sleep(0.2)
    os.system(str1 + 'shell input tap 235 355')  # 写入警告(汇报舰船)
    time.sleep(0.2)
    os.system(str1 + 'shell input tap 340 190')  # 点发送(退出更多)
    time.sleep(0.2)
    os.system(str1 + 'shell input tap 340 510')  # 发送
    time.sleep(1)
    os.system(str1 + 'shell input tap 774 502')  # 退出聊天区
    time.sleep(0.5)


def send_we_chat(tag, num):
    if wx_group_name == '':
        return
    mutex.acquire()
    send_msg(f'{path}/{tag}_{num}.png', msg_type=3)
    context = f"{tag} {wx_context}"
    send_msg(context, msg_type=1)
    mutex.release()


def listening(tag):
    # *截图->裁剪->识别->动作（游戏频道发送， 微信发送）

    def task2(tag):
        num = 0
        while True:
            screenc(tag, 1)
            # 检测舰船列表, 发送 微信
            time.sleep(0.5)
            crop(918, 44, 956, 153, f'{path}/{tag}_1.png', f'new_{tag}_list.png')
            i3, i4 = load_image(f"{path}/new_{tag}_list.png", f"{path}/tem/list.png")
            list_status, list_mac_v = if_img_i(i3, i4)

            if list_mac_v != 0.0 and list_mac_v < 0.10:
                if wx_group_name == '':
                    continue

                if num < 1:
                    num += 1
                    print('二次检测')
                    time.sleep(2)
                    continue
                # 防误报  二次检测

                num = 0
                print(tag + '检测到舰船列表有人', list_mac_v)
                send_we_chat(tag, 1)
                i1, i2 = load_image(f"{path}/new_{tag}_playerList.png", f"{path}/old_{tag}_playerList.png")
                cv2.imwrite(f'{path}/old_{tag}_playerList.png', i1, [cv2.IMWRITE_PNG_COMPRESSION, 0])
                time.sleep(30)  # 播报一次后延迟播报,避免信息轰炸
                continue

    t = threading.Thread(target=task2, args=(tag,))
    t.start()

    while True:
        screenc(tag, 2)
        time.sleep(con_val)
        # 第一次识别后, 判断是否检测舰船列表, 动作结束后将new playerList覆盖掉old
        time.sleep(0.35)
        crop(774, 502, 956, 537, f'{path}/{tag}_2.png', f'new_{tag}_playerList.png')
        i1, i2 = load_image(f"{path}/new_{tag}_playerList.png", f"{path}/old_{tag}_playerList.png")
        list_status, list_mac_v = if_img_i(i1, i2)

        # 疑似故障等待
        if list_mac_v <= 0.01:
            print(tag, '检测失败,疑似元素位置错误')
            time.sleep(3)
            continue

        # 检测到本地有红白, 发送游戏频道
        if list_status:
            print(tag + '警告')
            send_game_massage(tag)
            cv2.imwrite(f'{path}/old_{tag}_playerList.png', i1, [cv2.IMWRITE_PNG_COMPRESSION, 0])
            time.sleep(5)


start()
