# logic/word_to_hwp.py

import time
import subprocess
import win32gui
import win32con
import win32api
import win32com.client


def copy_to_hwp():
    try:
        subprocess.Popen("C:\\Program Files\\Windows NT\\Accessories\\wordpad.exe")
        time.sleep(2)

        wordpad_hwnd = win32gui.FindWindow(None, "문서 - 워드패드")
        if not wordpad_hwnd:
            raise Exception("WordPad 창을 찾을 수 없습니다.")

        win32gui.SetForegroundWindow(wordpad_hwnd)
        time.sleep(0.5)

        win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
        win32api.keybd_event(ord('V'), 0, 0, 0)
        win32api.keybd_event(ord('V'), 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(0.5)

        win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
        win32api.keybd_event(ord('A'), 0, 0, 0)
        win32api.keybd_event(ord('A'), 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(0.5)

        win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
        win32api.keybd_event(ord('C'), 0, 0, 0)
        win32api.keybd_event(ord('C'), 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(1)

        hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
        time.sleep(1)
        hwp.Run("Paste")

        subprocess.call("taskkill /f /im wordpad.exe", shell=True)
        subprocess.call("taskkill /f /im winword.exe", shell=True)
    except Exception as e:
        try:
            subprocess.call("taskkill /f /im wordpad.exe", shell=True)
            subprocess.call("taskkill /f /im winword.exe", shell=True)
        except:
            pass
        raise e
