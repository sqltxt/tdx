import win32gui
import win32api
from win32.lib import win32con
import time
import win32process
import datetime
import win32ui

path =  'C:/new_tdx/'
handle = win32process.CreateProcess(path+'TdxW.exe','',None,None,0,win32process.CREATE_NO_WINDOW,None,path,win32process.STARTUPINFO())#打开TB,获得其句柄
time.sleep(1)

TDX_handle = win32gui.FindWindow('#32770','通达信金融终端V7.35')
#print(hex(TDX_handle))
time.sleep(1)
Tab_handle = win32gui.FindWindowEx(TDX_handle,None,'SysTabControl32','Tab1')
#print(hex(Tab_handle))
time.sleep(1)
p=win32gui.GetWindowRect(Tab_handle)
#print(p)
#print(p[2])
#print(p[3])
win32api.SetCursorPos([p[0]+170,p[1]+7])
win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0,0,0)
time.sleep(2)
win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0,0,0)
time.sleep(2)
win32gui.PostMessage(win32gui.FindWindowEx(TDX_handle,None,'Button','登录'),win32con.BM_CLICK,0,0)
time.sleep(5)
win32gui.PostMessage(win32gui.FindWindow('#32770','通达信信息'),win32con.WM_CLOSE,0,0)
time.sleep(10)
win32api.keybd_event(17,0,0,0)#ctrl键位码是17
win32api.keybd_event(84,0,0,0)#t键位码是84
win32api.keybd_event(84,0,win32con.KEYEVENTF_KEYUP,0)#释放按键
win32api.keybd_event(17,0,win32con.KEYEVENTF_KEYUP,0)
time.sleep(5)

#win32gui.SendMessage(Tab_handle,0x130C,1,0)#尼玛，这个代码整整找了一天好吗！
tjxg = win32gui.FindWindow('#32770','条件选股')
gs = win32gui.FindWindowEx(win32gui.FindWindow('#32770','条件选股'),None,'ComboBox','UPN          - 连涨数天        ')
#print(hex(gs))
time.sleep(5)
win32gui.PostMessage(win32gui.FindWindowEx(win32gui.FindWindow('#32770','条件选股'),None,'Button','加入条件'),win32con.BM_CLICK,0,0)
time.sleep(3)
zs = win32gui.FindWindowEx(win32gui.FindWindow('#32770','条件选股'),None,'Button','执行选股')
#print(hex(zs))
win32gui.PostMessage(win32gui.FindWindowEx(win32gui.FindWindow('#32770','条件选股'),None,'Button','执行选股'),win32con.BM_CLICK,0,0)
time.sleep(40)
win32gui.PostMessage(win32gui.FindWindow('#32770','条件选股'),win32con.WM_CLOSE,0,0)
time.sleep(3)
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
#print(dhk.__getattribute_())
#win32gui.GetDlgItemText(dhk,1)
#win32gui.PostMessage(dhk,win32con.WM_COMMAND, win32gui.GetMenuItemID(win32gui.GetSubMenu(win32gui.GetMenu(dhk),6),23),0)#

win32api.keybd_event(18,0,0,0)#alt键位码是17
win32api.keybd_event(115,0,0,0)#F4键位码是115
win32api.keybd_event(115,0,win32con.KEYEVENTF_KEYUP,0)#释放按键
win32api.keybd_event(18,0,win32con.KEYEVENTF_KEYUP,0)

time.sleep(3)
#win32gui.PostMessage(win32gui.FindWindowEx(win32gui.FindWindow('#32770','退出确认'),None,'Button','退出'),win32con.BM_CLICK,0,0)
