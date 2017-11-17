import win32gui
import win32api
from win32.lib import win32con
import time
import win32process
import datetime
import win32ui
import ExportEBK
import ReadEBK
import sys
path =  'C:/new_tdx/'
Index = 86#系统选股公式一共83个
corp_id = 'wx87780fb826353ecc'
secret = 'XlOvsTpuNaINPsV2Wm6TMRLt_k1lgxZjJwrSaVND0Lo'
agentid = 1000003

#打开通达信
def Open_TDX():
    try:
        handle = win32process.CreateProcess(path+'TdxW.exe','',None,None,0,win32process.CREATE_NO_WINDOW,None,path,win32process.STARTUPINFO())#打开TB,获得其句柄
        time.sleep(3)
        TDX_handle = win32gui.FindWindow('#32770','通达信金融终端V7.35')
    except Exception as e:
        ReadEBK.wx_msg(corp_id, secret,agentid,sys._getframe().f_code.co_name+'\t'+str(e))
        
#免费行情
def Free_quotation():
    time.sleep(3)
    try:
        Tab_handle = win32gui.FindWindowEx(win32gui.FindWindow('#32770','通达信金融终端V7.35'),None,'SysTabControl32','Tab1')
        #print(hex(Tab_handle))
        #time.sleep(1)
        p=win32gui.GetWindowRect(Tab_handle)
        #print(p)
        #print(p[2])
        #print(p[3])
        win32api.SetCursorPos([p[0]+170,p[1]+7])
        time.sleep(1)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0,0,0)
        time.sleep(1)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0,0,0)
        time.sleep(1)
        win32gui.PostMessage(win32gui.FindWindowEx(win32gui.FindWindow('#32770','通达信金融终端V7.35'),None,'Button','登录'),win32con.BM_CLICK,0,0)
        time.sleep(3)
    except Exception as e:
        ReadEBK.wx_msg(corp_id, secret,agentid,sys._getframe().f_code.co_name+'\t'+str(e))

#登录信息框
def Info():
    try:
        win32gui.PostMessage(win32gui.FindWindow('#32770','通达信信息'),win32con.WM_CLOSE,0,0)
        time.sleep(5)
    except Exception as e:
        ReadEBK.wx_msg(corp_id, secret,agentid,sys._getframe().f_code.co_name+'\t'+str(e))
#条件选股 Ctrl+T
def Ctrl_T():
    try:
        win32api.keybd_event(17,0,0,0)#ctrl键位码是17
        win32api.keybd_event(84,0,0,0)#t键位码是84
        win32api.keybd_event(84,0,win32con.KEYEVENTF_KEYUP,0)#释放按键
        win32api.keybd_event(17,0,win32con.KEYEVENTF_KEYUP,0)
        time.sleep(3)
        #win32gui.SendMessage(Tab_handle,0x130C,1,0)#尼玛，这个代码整整找了一天好吗！
    except Exception as e:
        ReadEBK.wx_msg(corp_id, secret,agentid,sys._getframe().f_code.co_name+'\t'+str(e))

#按索引选择公式
def Stock_option(Index):
    try:
        #tjxg = win32gui.FindWindow('#32770','条件选股')
        gs = win32gui.FindWindowEx(win32gui.FindWindow('#32770','条件选股'),None,'ComboBox',None)
        #time.sleep(1)
        print(hex(gs))
        win32gui.SendMessage(gs,win32con.CB_SHOWDROPDOWN,1,0)  #展开ComboBox列表框
        time.sleep(1)
        win32gui.SendMessage(gs,win32con.CB_SETCURSEL,Index,0)#指向指定记录号
        time.sleep(1)
        win32gui.SendMessage(gs,win32con.WM_SETFOCUS,0,0)#选中按钮
        time.sleep(1)
        win32gui.SendMessage(gs,win32con.WM_KEYDOWN,0,0)#模拟按下指定键
        time.sleep(1)
        win32gui.SendMessage(gs,win32con.WM_KEYUP,0,0) 
        time.sleep(1)
    except Exception as e:
        ReadEBK.wx_msg(corp_id, secret,agentid,sys._getframe().f_code.co_name+'\t'+str(e))
#加入条件
def Join_condition():
    try:
        #time.sleep(5)
        win32gui.PostMessage(win32gui.FindWindowEx(win32gui.FindWindow('#32770','条件选股'),None,'Button','加入条件'),win32con.BM_CLICK,0,0)
        time.sleep(3)
    except Exception as e:
        ReadEBK.wx_msg(corp_id, secret,agentid,sys._getframe().f_code.co_name+'\t'+str(e))
#执行选股
def EXE():
    try:
        while win32gui.FindWindowEx(win32gui.FindWindow('#32770','条件选股'),None,'Static','选股完毕. ') == 0:
            zs = win32gui.FindWindowEx(win32gui.FindWindow('#32770','条件选股'),None,'Button','执行选股')
            print(zs)
            print(type(zs))
            win32gui.PostMessage(win32gui.FindWindowEx(win32gui.FindWindow('#32770','条件选股'),None,'Button','执行选股'),win32con.BM_CLICK,0,0)
            time.sleep(7)
            print("执行选股")
            break
    except Exception as e:
        ReadEBK.wx_msg(corp_id, secret,agentid,sys._getframe().f_code.co_name+'\t'+str(e))
#补数据
def Complement_data():
    try:
        time.sleep(3)
        if win32gui.FindWindowEx(win32gui.FindWindow('#32770','TdxW'),None,'Button','是(&Y)') != 0:
            win32gui.PostMessage(win32gui.FindWindowEx(win32gui.FindWindow('#32770','TdxW'),None,'Button','是(&Y)'),win32con.BM_CLICK,0,0)
            time.sleep(5)
            print("补数据")
    except Exception as e:
        ReadEBK.wx_msg(corp_id, secret,agentid,sys._getframe().f_code.co_name+'\t'+str(e))
#关闭选股器
def Close_Stockpickingmachine():
    time.sleep(5)
    try:
        #while win32gui.FindWindowEx(win32gui.FindWindow('#32770','条件选股'),None,'Static','选股完毕. ') == 0:
            print("进入循环")
            time.sleep(1)
            #if win32gui.FindWindowEx(win32gui.FindWindow('#32770','条件选股'),None,'Static','选股完毕. ') != 0:
            win32gui.PostMessage(win32gui.FindWindow('#32770','条件选股'),win32con.WM_CLOSE,0,0)
            time.sleep(3)
            print("选股完毕")
                #break
    except Exception as e:
        ReadEBK.wx_msg(corp_id, secret,agentid,sys._getframe().f_code.co_name+'\t'+str(e))
def Close_TDX():
    try:
        time.sleep(3)
        lstj=win32gui.FindWindow('TdxW_MainFrame_Class','通达信金融终端V7.35 - [行情表-临时条件股]')
        print(hex(lstj))
        win32gui.PostMessage(lstj,win32con.WM_CLOSE,0,0)
        time.sleep(3)
        win32gui.PostMessage(win32gui.FindWindowEx(win32gui.FindWindow('#32770','退出确认'),None,'Button','退出'),win32con.BM_CLICK,0,0)
        time.sleep(3)
    except Exception as e:
        ReadEBK.wx_msg(corp_id, secret,agentid,sys._getframe().f_code.co_name+'\t'+str(e))
        
if __name__=='__main__':
    while 1:
        if int(time.strftime("%H%M%S"))== 90000:
            pass
            #Open_TDX()
        elif int(time.strftime("%H%M%S"))== 121400:#or int(time.strftime("%H%M%S"))== 114200:
            Open_TDX()
            Free_quotation()
            Info()
            Ctrl_T()
            Stock_option(Index)
            Join_condition()
            EXE()
            Complement_data()
            Close_Stockpickingmachine()
            ExportEBK.Sequential_execution()
        elif int(time.strftime("%H%M%S"))== 123000:
            ReadEBK.Send_wx()
        elif int(time.strftime("%H%M%S"))== 160000:
            Close_TDX()
        time.sleep(1)
