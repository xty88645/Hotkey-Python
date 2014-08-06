# -*- coding: utf-8 -*-
#
# Smurfs V1.1.2
# by Alex Maven
# 2014/2/11
#
# 托盘显示

import win32api, win32gui
import win32con, winerror
import os
import subprocess
import pyHook
import ConfigParser
import itertools
import sys

# 全局变量
Switch_Button = 1  # 程序控制开关
arr_Temp = []  # 存储临时按键的数组

# 读取配置文件
config = ConfigParser.ConfigParser()                
config.readfp(open('UserMsg.ini'))
primerKey = config.get("userInit", "primerKey")  # 读取关键按键，开始匹配自定义程序
ClickNum = int(config.get("userInit", "ClickNum"))  # 读取关键按键需要按几次激活
Hotkey_With2 = []  # 全部2个快捷键的程序快捷键配置
Hotkey_With3 = []  # 全部3个快捷键的程序快捷键配置

Switch_Button = 1  # 程序控制开关

def initMain():
    # 初始化快捷键配置
    
    global config
    ReadIni = config.sections() 
    ReadIni.remove('userInit')
    for i in ReadIni:
        if len(config.get(i, "Hotkeys")) == 2:
            Hotkey_With2.append([config.get(i, "FileLocate"), config.get(i, "Hotkeys")])
        elif len(config.get(i, "Hotkeys")) == 3:
            Hotkey_With3.append([config.get(i, "FileLocate"), config.get(i, "Hotkeys")])
    print("Initialization configured successfully!")
    print("-----------------------------------------------")

def onKeyboardEvent(event):
    # 监听键盘事件
    
    global Switch_Button, primerKey, ClickNum, Hotkey_With2, Hotkey_With3, arr_Temp
    counts = 0
    
    if Switch_Button > 0:  # 检测到关键按键触发    
        arr_Temp.append(event.Key)
        if len(arr_Temp) == ClickNum:            
            for ix in arr_Temp:
                if ix == primerKey:
                    counts += 1
                    
            if counts == ClickNum:
                Switch_Button = -1
                initMain()
                counts = 0
                arr_Temp = []
                print("Waiting for the start command!")
                print("-----------------------------------------------")
                return
            else:
                arr_Temp = []
                
    if Switch_Button < 0:  # 开始记录用户键盘录入与配置文件进行匹配
        arr_Temp.append(event.Key)
        if arr_Temp.count(primerKey) == ClickNum:
            arr_Temp = []
            print("Waiting for the start command!")
            print("-----------------------------------------------")
        else:
            if len(arr_Temp) == 2:
                arrTemp = "".join(itertools.chain(*arr_Temp))
                for a, b in Hotkey_With2:             
                    if b == arrTemp:
                        Switch_Button = 1
                        subprocess.Popen(a, shell=True)
                        print("Program started successfully!")
                        print("-----------------------------------------------")
                        arr_Temp = []
                        arrTemp = "" 
                         
            elif len(arr_Temp) == 3:
                arrTemp = "".join(itertools.chain(*arr_Temp))
                for a, b in Hotkey_With3:
                    if b == arrTemp:
                        Switch_Button = 1
                        subprocess.Popen(a, shell=True)
                        print("Program started successfully!")
                        print("-----------------------------------------------")
                        arr_Temp = []
                        arrTemp = ""

    # 监听函数的返回值
    return True

def mainStart():

    hm = pyHook.HookManager()  # 创建一个“钩子”管理对象
    hm.KeyDown = onKeyboardEvent  # 监听所有键盘事件
    hm.HookKeyboard()  # 设置键盘“钩子”
    win32gui.PumpMessages()  # 进入循环监听状态
    
def mainStop():
    sys.exit(0)

class MainWindow:
    def __init__(self):
        msg_TaskbarRestart = win32gui.RegisterWindowMessage("TaskbarCreated");
        message_map = {
                msg_TaskbarRestart: self.OnRestart,
                win32con.WM_DESTROY: self.OnDestroy,
                win32con.WM_COMMAND: self.OnCommand,
                win32con.WM_USER + 20 : self.OnTaskbarNotify,
        }
        # Register the Window class.
        wc = win32gui.WNDCLASS()
        hinst = wc.hInstance = win32api.GetModuleHandle(None)
        wc.lpszClassName = "PythonTaskbarDemo"
        wc.style = win32con.CS_VREDRAW | win32con.CS_HREDRAW;
        wc.hCursor = win32api.LoadCursor(0, win32con.IDC_ARROW)
        wc.hbrBackground = win32con.COLOR_WINDOW
        wc.lpfnWndProc = message_map  # could also specify a wndproc.

        # Don't blow up if class already registered to make testing easier
        try:
            win32gui.RegisterClass(wc)
        except win32gui.error, err_info:
            if err_info.winerror != winerror.ERROR_CLASS_ALREADY_EXISTS:
                raise

        # Create the Window.
        style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
        self.hwnd = win32gui.CreateWindow(wc.lpszClassName, "Taskbar Demo", style, \
                0, 0, win32con.CW_USEDEFAULT, win32con.CW_USEDEFAULT, \
                0, 0, hinst, None)
        win32gui.UpdateWindow(self.hwnd)
        self._DoCreateIcons()
    def _DoCreateIcons(self):
        # Try and find a custom icon
        hinst = win32api.GetModuleHandle(None)
        iconPathName = "Smurfs.ico"
        if os.path.isfile(iconPathName):
            icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
            hicon = win32gui.LoadImage(hinst, iconPathName, win32con.IMAGE_ICON, 0, 0, icon_flags)
        else:
            print "Can't find a Python icon file - using default"
            hicon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)

        flags = win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP
        nid = (self.hwnd, 0, flags, win32con.WM_USER + 20, hicon, "Smurfs v1.1.2")
        try:
            win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, nid)
        except win32gui.error:
            print "Failed to add the taskbar icon - is explorer running?"

    def OnRestart(self, hwnd, msg, wparam, lparam):
        self._DoCreateIcons()

    def OnDestroy(self, hwnd, msg, wparam, lparam):
        nid = (self.hwnd, 0)
        win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, nid)
        win32gui.PostQuitMessage(0)  # Terminate the app.

    def OnTaskbarNotify(self, hwnd, msg, wparam, lparam):
        global Switch_Button
        if lparam == win32con.WM_LBUTTONUP:
            if Switch_Button > 0:
                ct = win32api.GetConsoleTitle()   
                hd = win32gui.FindWindow(0, ct)   
                win32gui.ShowWindow(hd, 0)
                Switch_Button = -1
            elif Switch_Button < 0:
                ct = win32api.GetConsoleTitle()   
                hd = win32gui.FindWindow(0, ct)   
                win32gui.ShowWindow(hd, 1)
                Switch_Button = 1
        elif lparam == win32con.WM_RBUTTONUP:
            menu = win32gui.CreatePopupMenu()
            win32gui.AppendMenu(menu, win32con.MF_STRING, 1023, "Start")
            win32gui.AppendMenu(menu, win32con.MF_STRING, 1024, "Configure")
            win32gui.AppendMenu(menu, win32con.MF_STRING, 1025, "Exit")
            pos = win32gui.GetCursorPos()
            win32gui.SetForegroundWindow(self.hwnd)
            win32gui.TrackPopupMenu(menu, win32con.TPM_LEFTALIGN, pos[0], pos[1], 0, self.hwnd, None)
            win32gui.PostMessage(self.hwnd, win32con.WM_NULL, 0, 0)
        return 1

    def OnCommand(self, hwnd, msg, wparam, lparam):
        uin = win32api.LOWORD(wparam)
        if uin == 1023:
            mainStart()
        elif uin == 1024:
            subprocess.Popen("UserMsg.ini", shell=True)
        elif uin == 1025:
            print "Goodbye"
            win32gui.DestroyWindow(self.hwnd)
            mainStop()
        else:
            print "Unknown command -", uin

def main():
    MainWindow()
    win32gui.PumpMessages()

if __name__ == '__main__':
    # 程序信息
    print("-----------------------------------------------")
    print("|                Smurfs V1.1.2                |")
    print("|                by Alex Maven                |")
    print("|                2014.2.11                    |")
    print("-----------------------------------------------\n")
    print("Program started successfully!")
    print("-----------------------------------------------")
    
    main()
