import os
import win32gui, win32ui, win32con, win32api
import time
import cv2 as cv
import numpy as np
from pymouse import PyMouse
from pykeyboard import PyKeyboard


skip_special = False
scale = 1.25        # 当Windows系统显示缩放时需要调整scale至缩放比例
send_clipborad_confirm = False

# win32api 截图，较快
def window_capture(filename):
  hwnd = 0 # 窗口的编号，0号表示当前活跃窗口
  # 根据窗口句柄获取窗口的设备上下文DC（Divice Context）
  hwndDC = win32gui.GetWindowDC(hwnd)
  # 根据窗口的DC获取mfcDC
  mfcDC = win32ui.CreateDCFromHandle(hwndDC)
  # mfcDC创建可兼容的DC
  saveDC = mfcDC.CreateCompatibleDC()
  # 创建bigmap准备保存图片
  saveBitMap = win32ui.CreateBitmap()
  # 获取监控器信息
  MoniterDev = win32api.EnumDisplayMonitors(None, None)
  w = MoniterDev[0][2][2]
  h = MoniterDev[0][2][3]
  # print w,h　　　#图片大小
  # 为bitmap开辟空间
  saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)
  # 高度saveDC，将截图保存到saveBitmap中
  saveDC.SelectObject(saveBitMap)
  # 截取从左上角（0，0）长宽为（w，h）的图片
  saveDC.BitBlt((0, 0), (w, h), mfcDC, (0, 0), win32con.SRCCOPY)
  saveBitMap.SaveBitmapFile(saveDC, filename)

# 找到搜索框返回点击位置
def find_icon(icon_filename="data/search.png"):
    window_capture("data/temp.png")
    capture = cv.imread("data/temp.png")
    template = cv.imread(icon_filename)
    h, w = template.shape[:2]

    # 模板匹配
    res = cv.matchTemplate(capture, template, cv.TM_CCOEFF_NORMED)
    max_corr = np.max(res)
    if max_corr <0.95:
        return -1, -1
    left_up = np.where(res == max_corr)[::-1]
    right_bottom = (left_up[0] + w, left_up[1] + h)
    cv.rectangle(capture, left_up, right_bottom, (0, 0, 255), 2)
    cv.imshow("capture", capture)
    cv.waitKey(500)
    cv.destroyAllWindows()
    click_x = int((left_up[0]+right_bottom[0])/2)
    click_y = int((left_up[1]+right_bottom[1])/2)
    return click_x, click_y

# win32api单击
# 暂未使用
def click(x, y):
    win32api.SetCursorPos([x, y])
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)


if __name__ == "__main__":
    k = PyKeyboard()
    m = PyMouse()
    # 找到搜索框位置
    search_icon_x, search_icon_y = find_icon()
    if search_icon_x == -1:
        raise RuntimeError("请打开tim窗口")
    # 读取qq号们
    f = open("data/numbers.txt", 'r', encoding="utf-8")
    numbers = f.readlines()

    for line in numbers:
        if(line[0]=="#"):
            continue
        if (line == "special\n"):
            break
        number, name = line.split(":")
        # 点搜索框
        m.click(int(search_icon_x/scale), int(search_icon_y/scale))
        time.sleep(0.5)
        # 输入qq号
        k.type_string(number)
        time.sleep(0.5)
        # 按回车
        k.tap_key(k.enter_key)
        time.sleep(0.5)
        # 发送消息内容(来自剪贴板)
        if send_clipborad_confirm:
            k.press_keys([k.control_key, 'v'])
            k.tap_key(k.enter_key)
        time.sleep(1)
        print(name.replace("\n", "完成"))
