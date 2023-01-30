import win32gui,win32process
import win32con
import subprocess
import keyboard
import time
import webbrowser
from win32gui import GetWindowText, GetForegroundWindow
import psutil
from win32com.client import GetObject

def meetattach(f=False):
    attached = False
    while not attached:
        fwindow = GetWindowText(GetForegroundWindow())
        if "Meet" in fwindow and fwindow.count("–") == 1:
            global hwnd
            global meetname
            global scode
            meetname = fwindow
            scode = meetname[7:-16]
            hwnd = win32gui.FindWindow(None, meetname)
            attached = True
            if f:
                print("Meet Trouvé!")
                print(f"Accroché au Meet {scode}")

def getpid():
    global pid
    pid = 0
    threadid,pid = win32process.GetWindowThreadProcessId(hwnd)

def meetcheck():
    check = 0
    check = win32gui.FindWindow(None, meetname)
    if check != 0:
        return True
    else:
        return False

def chromecheck():
    fwindow = GetWindowText(GetForegroundWindow())
    if "Google Chrome" in fwindow:
        return True
    else:
        return False

def pid_exists(pid):
        import ctypes
        kernel32 = ctypes.windll.kernel32
        SYNCHRONIZE = 0x100000

        process = kernel32.OpenProcess(SYNCHRONIZE, 0, pid)
        if process != 0:
            kernel32.CloseHandle(process)
            return True
        else:
            return False

meetattach(True)

while True:
    getpid()
    if not pid_exists(pid):
        webbrowser.open(f'https://meet.google.com/{scode}', new=2)
        pid = 0
        meetattach()
        getpid()
        time.sleep(5)

    if not meetcheck():
        keyboard.press_and_release('ctrl+w')
        time.sleep(1)

    elif not chromecheck():
        try:
            hwnd = win32gui.FindWindow(None, meetname)
            win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
            win32gui.SetForegroundWindow(hwnd)
        except:
            continue
        time.sleep(2)
