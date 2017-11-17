import win32gui
import win32api
from win32.lib import win32con
import time
import win32process
import datetime
import win32ui
from commctrl import LVM_GETITEMTEXT, LVM_GETITEMCOUNT

def handle_window(hwnd,extra):#TB_handle句柄
    if win32gui.IsWindowVisible(hwnd):
        if extra in win32gui.GetWindowText(hwnd):
            global handle
            handle= hwnd

lstj=win32gui.FindWindow('TdxW_MainFrame_Class','通达信金融终端V7.35 - [行情表-临时条件股]')
print(hex(lstj))
#win32gui.PostMessage(lstj,win32con.WM_CLOSE,0,0)
dhk = win32gui.FindWindowEx(lstj,None,'#32770',None)
print(hex(dhk))
print(dir(dhk))

pdhk=win32gui.GetWindowRect(dhk)
win32api.SetCursorPos([pdhk[0]+330,pdhk[1]+10])
time.sleep(3)
win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0,0,0)
time.sleep(2)
win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0,0,0)
time.sleep(2)
win32api.SetCursorPos([pdhk[0]+380,pdhk[1]+20+480])
time.sleep(3)
win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0,0,0)
time.sleep(2)
win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0,0,0)
time.sleep(2)
zdk=[]
dc = win32gui.SendMessage(win32gui.FindWindowEx(win32gui.FindWindowEx(win32gui.FindWindow('#32770','自定义板块设置'),None,'#32770','自定义板块管理'),None,'SysListView32','CFQS'),LVM_GETITEMCOUNT)
zdybkgl=win32gui.FindWindowEx(win32gui.FindWindowEx(win32gui.FindWindow('#32770','自定义板块设置'),None,'#32770','自定义板块管理'),None,'SysListView32','CFQS')
e = win32gui.GetWindowRect (zdybkgl)
print(e)
win32api.SetCursorPos([e[0]+50,e[1]+25])
time.sleep(3)
win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0,0,0)
time.sleep(2)
win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0,0,0)
time.sleep(2)
dcbk = win32gui.FindWindowEx(win32gui.FindWindowEx(win32gui.FindWindow('#32770','自定义板块设置'),None,'#32770','自定义板块管理'),None,'Button','导出板块')
print(hex(dcbk))
win32gui.PostMessage(dcbk,win32con.BM_CLICK,0,0)
time.sleep(2)
wjm = win32gui.FindWindowEx(win32gui.FindWindow('#32770','另存为'),None,'Edit',None)
print(hex(wjm))
win32gui.SendMessage(wjm,win32con.WM_SETTEXT,0,'每日选股'+str(int(time.strftime("%Y%m%d")))+'.EBK')
time.sleep(2)
bc = win32gui.FindWindowEx(win32gui.FindWindow('#32770','另存为'),None,'Button','保存(&S)')
print(hex(bc))

win32gui.PostMessage(bc,win32con.BM_CLICK,0,0)
time.sleep(2)
qrlcw = win32gui.FindWindow('#32770','确认另存为')
print(hex(qrlcw))
global handle
handle = None
win32gui.EnumChildWindows(qrlcw,handle_window,'是(&Y)')
if handle!= None:
    win32gui.PostMessage(handle,win32con.BM_CLICK,0,0)
time.sleep(2)
win32gui.PostMessage(win32gui.FindWindowEx(win32gui.FindWindow('#32770','TdxW'),None,'Button','确定'),win32con.BM_CLICK,0,0)
time.sleep(2)
qx = win32gui.FindWindowEx(win32gui.FindWindow('#32770','自定义板块设置'),None,'Button','取消')
print(hex(qx))
win32gui.PostMessage(qx,win32con.BM_CLICK,0,0)

