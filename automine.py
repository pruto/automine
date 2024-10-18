import pyautogui
import numpy
import cv2
import sys
import win32gui
import win32api
import win32con
import win32com.client
from time import sleep
from math import floor

handle = None

def click(bt, num=0):
    sleep(.03)
    if num == 0:
        pyautogui.moveTo(bt[0], bt[1], duration=.0)
    else:
        pyautogui.moveTo(bt[0]+bt[2]*(num%bt[4]), bt[1]+bt[3]*floor(num/bt[4]), duration=.0)
    sleep(.03)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    sleep(.03)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)

def test_pixel(clrs, bt, num=0):
    sleep(.15)
    if num == 0:
        px = bt[0]
        py = bt[1]
    else:
        px = bt[0]+bt[2]*(num%bt[4])
        py = bt[1]+bt[3]*floor(num/bt[4])
    ret = False
    for i in clrs:
        if pyautogui.pixelMatchesColor(int(px), int(py), i):
            ret = True
            break
    return ret

def test_abort():
    global handle
    if handle != win32gui.GetForegroundWindow():
        sys.exit()

def wait_click(dur, bt, num=0):
    while dur >= 0:
        sleep(.125)
        dur -= .125
        test_abort()
    click(bt, num)

def main():
    # 获取窗口并挪到前台
    global handle
    handle = win32gui.FindWindow(None, '地下城与勇士：创新世纪')
    win32gui.ShowWindow(handle, win32con.SW_NORMAL)
    win32gui.BringWindowToTop(handle)
    shell = win32com.client.Dispatch("WScript.Shell")
    shell.SendKeys('%')
    win32gui.SetForegroundWindow(handle)

    # 常量
    winleft, wintop, winright, winbottom = win32gui.GetWindowRect(handle)
    winw = winright-winleft
    winh = winbottom-wintop
    wincx = winleft+.5*winw
    wincy = wintop+.5*winh

    # posx, posy, diffx, diffy, cntx, total
    btcollect = [wincx+264, wincy+164, ]
    btfacs = [wincx-234, wincy-120, 120, 180, 5, 10]
    btfacsclose = [wincx+120, wincy-116]
    btfacsunlock = [wincx+38, wincy+66]
    btmap = [wincx+78, wincy+150]
    btmapclose = [wincx+120, wincy-116]
    btmng = [wincx+148, wincy+150]
    btmngclose = [wincx+243, wincy-155]
    btmngleft = [wincx-110, wincy-130]
    btmngright = [wincx+120, wincy-130]
    btfacup = [wincx+36, wincy+160]
    btrobots = [wincx-179, wincy-72, 56, 76, 8, 24]
    btrobotup = [wincx+46, wincy+26]
    btrobotclose = [wincx+122, wincy-80]

    colorenabled = [(6, 153, 184), (9, 159, 194)]
    colordisabled = [(93, 94, 98), (96, 99, 107)]
    colornorobot = [(6, 153, 180), (7, 154, 183)]

    while True:
        for ifac in range(btfacs[5]):
            wait_click(1, btcollect)
            wait_click(2, btfacs, ifac)
            if test_pixel(colordisabled, btfacsunlock):
                wait_click(1, btfacsclose)
                break
            elif test_pixel(colorenabled, btfacsunlock):
                wait_click(1, btfacsunlock)
                wait_click(1, btfacs, ifac)
            wait_click(5, btmng)
            wait_click(1, btmngleft)
            if test_pixel(colorenabled, btfacup):
                wait_click(1, btfacup)
            wait_click(1, btmngright)
            for irobot in range(btrobots[5]):
                if test_pixel(colornorobot, btrobots, irobot):
                    break
                wait_click(.5, btrobots, irobot)
                if test_pixel(colorenabled, btrobotup):
                    wait_click(.5, btrobotup)
                wait_click(1, btrobotclose)
            wait_click(1, btmngclose)
            wait_click(1, btmap)

if __name__ == '__main__':
    main()
