import win32gui
import win32api
from win32.lib import win32con
import time
import win32process
import datetime
import win32ui
from commctrl import LVM_GETITEMTEXT, LVM_GETITEMCOUNT
import ReadEBK
import sys
corp_id = 'wx87780fb826353ecc'
secret = 'XlOvsTpuNaINPsV2Wm6TMRLt_k1lgxZjJwrSaVND0Lo'
agentid = 1000003

def handle_window(hwnd,extra):#TB_handle句柄
    if win32gui.IsWindowVisible(hwnd):
        if extra in win32gui.GetWindowText(hwnd):
            global handle
            handle= hwnd

def Menu():
    try:
        #工具
        pdhk=win32gui.GetWindowRect(win32gui.FindWindowEx(win32gui.FindWindow('TdxW_MainFrame_Class','通达信金融终端V7.35 - [行情表-临时条件股]'),None,'#32770',None))
        win32api.SetCursorPos([pdhk[0]+330,pdhk[1]+10])
        time.sleep(1)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0,0,0)
        time.sleep(1)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0,0,0)
        time.sleep(1)
        #自定义板块
        win32api.SetCursorPos([pdhk[0]+380,pdhk[1]+20+480])
        time.sleep(1)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0,0,0)
        time.sleep(1)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0,0,0)
        time.sleep(1)
    except Exception as e:
        ReadEBK.wx_msg(corp_id, secret,agentid,sys._getframe().f_code.co_name+'\t'+str(e))
        
#自定义板块管理->临时条件股
def CFQS():
    try:
        e = win32gui.GetWindowRect (win32gui.FindWindowEx(win32gui.FindWindowEx(win32gui.FindWindow('#32770','自定义板块设置'),None,'#32770','自定义板块管理'),None,'SysListView32','CFQS'))
        win32api.SetCursorPos([e[0]+50,e[1]+25])
        time.sleep(1)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0,0,0)
        time.sleep(1)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0,0,0)
        time.sleep(1)
    except Exception as e:
        ReadEBK.wx_msg(corp_id, secret,agentid,sys._getframe().f_code.co_name+'\t'+str(e))
        
#导出板块
def Export_plate():
    try:
        win32gui.PostMessage(win32gui.FindWindowEx(win32gui.FindWindowEx(win32gui.FindWindow('#32770','自定义板块设置'),None,'#32770','自定义板块管理'),None,'Button','导出板块'),win32con.BM_CLICK,0,0)
        time.sleep(1)
    except Exception as e:
        ReadEBK.wx_msg(corp_id, secret,agentid,sys._getframe().f_code.co_name+'\t'+str(e))
        
#文件名
def File_name():
    try:
        win32gui.SendMessage(win32gui.FindWindowEx(win32gui.FindWindow('#32770','另存为'),None,'Edit',None),win32con.WM_SETTEXT,0,'每日选股'+str(int(time.strftime("%Y%m%d")))+'.EBK')
        time.sleep(1)
    except Exception as e:
        ReadEBK.wx_msg(corp_id, secret,agentid,sys._getframe().f_code.co_name+'\t'+str(e))
        
#保存
def Save():
    try:
        win32gui.PostMessage(win32gui.FindWindowEx(win32gui.FindWindow('#32770','另存为'),None,'Button','保存(&S)'),win32con.BM_CLICK,0,0)
        time.sleep(1)
        global handle
        handle = None
        win32gui.EnumChildWindows(win32gui.FindWindow('#32770','确认另存为'),handle_window,'是(&Y)')
        if handle!= None:
            win32gui.PostMessage(handle,win32con.BM_CLICK,0,0)
        time.sleep(3)
    except Exception as e:
        ReadEBK.wx_msg(corp_id, secret,agentid,sys._getframe().f_code.co_name+'\t'+str(e))

def IsOK():
    try:
        #确定
        win32gui.PostMessage(win32gui.FindWindowEx(win32gui.FindWindow('#32770','TdxW'),None,'Button','确定'),win32con.BM_CLICK,0,0)
        time.sleep(1)
        #取消自定义板块设置
        win32gui.PostMessage(win32gui.FindWindowEx(win32gui.FindWindow('#32770','自定义板块设置'),None,'Button','取消'),win32con.BM_CLICK,0,0)
    except Exception as e:
        ReadEBK.wx_msg(corp_id, secret,agentid,sys._getframe().f_code.co_name+'\t'+str(e))

#顺序执行供外调
def Sequential_execution():
    time.sleep(5)
    Menu()
    CFQS()
    Export_plate()
    File_name()
    Save()
    IsOK()
