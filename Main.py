import win32gui
import win32con
import win32api
import pythoncom
import win32com
import time
import win32clipboard as w
from win32com import client
import PySimpleGUI as sg
import os
import webbrowser
import psutil

sg.theme('BrownBlue')
global pid
pid = os.getpid()


def create_txt(filename):
    filepath = os.getcwd()
    print(filepath)
    file_path = filepath + '\\' + filename + '.txt'
    print(file_path)
    #exit()
    file = open(file_path, 'w')
    file.close()


create_txt('QQdata')
with open(r'QQdata.txt', 'a+', encoding='utf-8') as test:
    test.truncate(0)

con = os.path.exists('contects.txt')
if con == True:
    with open('contects.txt', 'r', encoding='utf-8') as f:
        da = str(f.readlines())
        for daa in da:
            da = da.strip('\n')
            print(da)


elif con == False:
    create_txt('contects')
    with open(r'contects.txt', 'a+', encoding='utf-8') as test:
        test.truncate(0)

warr = ['------------[此文件将被覆写,若需要，请及时另存]------------']
with open("QQdata.txt", "w") as f:
    f.writelines(warr)


def setText(massage):
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_UNICODETEXT, massage)
    w.CloseClipboard()


def send_qq(massage):
    setText(massage)
    hwnd_title = dict()

    def get_all_hwnd(hwnd, mouse):
        if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd):
            hwnd_title.update({hwnd: win32gui.GetWindowText(hwnd)})

    win32gui.EnumWindows(get_all_hwnd, 0)
    for h, t in hwnd_title.items():
        if t != "":
            hwnd = win32gui.FindWindow('TXGuiFoundation', t)  # 获取qq窗口句柄
            if hwnd != 0:
                win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
                time.sleep(0.1)
                win32gui.SetForegroundWindow(hwnd)
                win32gui.SetActiveWindow(hwnd)
                time.sleep(0.1)
                win32gui.SendMessage(hwnd, 770, 0, 0)  #将剪贴板文本发送到QQ窗体
                win32gui.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
                win32gui.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)


def get_contects(mod):
    global name_list
    name_list = []
    ProData = ''
    if mod == 'A':
        l = ''
        while l != 'exit':
            global qqm
            qqm = [
                [sg.Text('输入最少一个QQ号按下按键以继续')],
                [sg.Text('支持同群非好友')],
                [sg.Text('确保已打开并登录‘QQ’客户端')],
                [sg.InputText('')],
                [sg.Text('系统内存占用:' + str(mem_percent))],
                [sg.Button('爷要输入这条qq号!', size=16), sg.Button('爷输入好了!', size=16),
                 sg.Button('关闭', size=16)], ]
            window = sg.Window('AmberQQ_Send', qqm, icon='favicon.ico')
            event, values = window.read()
            if event in (None, '关闭'):
                window.close();
                del window
                quit()
            elif event in ('核对qq号'):
                sg.popup_scrolled('AmberQQ_Send已输入QQ号', name_list)
            elif event in ('打开txt文件'):
                os.startfile(r'QQdata.txt')
            elif event in ('爷输入好了!'):
                l = 'exit'
            elif event in ('爷要输入这条qq号!'):
                global a
                a = values
            else:
                window.close();
                del window
                quit()
            window.close();
            del window
            qqm = [
                [sg.Text('输入最少一个QQ号按下按键以继续')],
                [sg.Text('支持同群非好友')],
                [sg.Text('确保已打开并登录‘QQ’客户端')],
                [sg.InputText('')],
                [sg.Text('系统内存占用:' + str(mem_percent))],
                [sg.Button('爷要输入这条qq号!', size=16), sg.Button('爷输入好了!', size=16),
                 sg.Button('关闭', size=16)], ]
            window = sg.Window('AmberQQ_Send', qqm, icon='favicon.ico')
            b = (list(a.values()))
            i = (",".join(b))
            if i != 'exit':
                PorData = name_list
                ProData = str(PorData)
                name_list.append(i)
                data = name_list
                with open("QQdata.txt", "w") as f:
                    for i in data:
                        i = str(i).strip('[').strip(']').replace(',', '').replace('\'', '') + '\n'
                        f.write(i)
        window.close();
        del window


def open_windows():
    qq_hwnd = win32gui.FindWindow(None, 'QQ')
    win32gui.ShowWindow(qq_hwnd, win32con.SW_SHOW)
    time.sleep(1)
    for i in name_list:
        massage = i
        setText(massage)
        pythoncom.CoInitialize()
        shell = win32com.client.Dispatch("WScript.Shell")
        pythoncom.CoInitialize()
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys('%')
        win32gui.SetForegroundWindow(qq_hwnd)
        win32gui.SetActiveWindow(qq_hwnd)
        time.sleep(1)
        win32gui.SendMessage(qq_hwnd, 770, 0, 0)
        time.sleep(1)
        win32gui.SetForegroundWindow(qq_hwnd)
        win32gui.SetActiveWindow(qq_hwnd)
        win32api.keybd_event(0x0D, win32api.MapVirtualKey(0x0D, 0), 0, 0)
        win32api.keybd_event(0x0D, win32api.MapVirtualKey(0x0D, 0), win32con.KEYEVENTF_KEYUP, 0)


key = 'R'
while key == 'R':
    mode = 'E'
    mem_percent = psutil.Process(pid).memory_percent()
    while (mode != 'A') and (mode != 'B'):
        start = [
            [sg.Text('支持同群非好友')],
            [sg.Text('确保已打开并登录‘QQ’客户端')],
            [sg.Text('确保已关闭QQ合并窗口')],
            [sg.Text('自动运行时请不要操控电脑!!!')],
            [sg.Text('本软件占用内存大小:' + str(mem_percent))],
            [sg.MLine(
                default_text='AmberQQ_Send 2.10版本' + '\n' + '新添功能↓↓↓:' + '\n' + '查看已输入qq号 \n 一键查看教程  \n 没什么卵用的小功能 \n 软件开发者:Amber \n Tim/QQ:2662344413 \n 对本软件有任何想法或建议的,欢迎联系我！！！',
                size=(50, 3)),
             [sg.Button('俺要看教程'), sg.Button('直接来罢!'), sg.Button('关闭')]
             ]]

        window = sg.Window('AmberQQ_Send', start, icon='favicon.ico')
        event, values = window.read()
        if event in (None, '关闭'):
            window.close();
            del window
            quit()

        elif event in ('直接来罢!'):
            window.close()
            mode = 'A'
        else:
            window.close()
            webbrowser.open(url='http://send.aerosotu.cn/', new=0, autoraise=True)
            window.close()
    get_contects(mode)
    open_windows()
    massag = [[sg.Text('输入要发送的消息(每人发送一次)')],
              [sg.Text('不过呢，在此之前你可以仔细核对你输入的qq号!')],
              [sg.Text('可以选择‘核对qq号’或者‘打开txt文件’')],
              [sg.InputText()],
              [sg.Button('发送', size=25), sg.Button('关闭', size=25)],
              [sg.Button('核对qq号', size=25), sg.Button('打开txt文件', size=25)]]
    window = sg.Window('AmberQQ_Send', massag, icon='favicon.ico')
    while True:
        event, values = window.read()
        if event in (None, '关闭'):
            window.close();
            del window
            break
        elif event in ('核对qq号'):
            sg.popup_scrolled('AmberQQ_Send已输入QQ号', name_list)
        elif event in ('打开txt文件'):
            os.startfile(r'QQdata.txt')
        else:
            window.close()
            massage = values[0]
    ct = int('1')
    send_qq(massage)
    ct = ct - 1
    time.sleep(0.1)
w.EmptyClipboard()
window.close()
