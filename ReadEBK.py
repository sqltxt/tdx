import time
import queue
import requests
import json

def wx_msg(corp_id, secret,agentid,msg):
    try:
        values = {'corpid' :corp_id,
              'corpsecret':secret
              }
        req = requests.post('https://qyapi.weixin.qq.com/cgi-bin/gettoken',params=values)
        token = json.loads(req.text)["access_token"]
        #try:
        dict_arr = {"touser": "@all",
                    "toparty": "@all",
                    "msgtype": "text",
                    "agentid": agentid,
                    "text": {"content": msg},
                    "safe": "0"}
        data = json.dumps(dict_arr,ensure_ascii=False,indent=2,sort_keys=True).encode('utf-8')
        reqs = requests.post("https://qyapi.weixin.qq.com/cgi-bin/message/send?access_token="+token,data)
    except Exception as e:
        with open('微信发送失败:'+str(int(time.strftime("%H%M%S")))+'.log') as f:
            f.write(e)
def Send_wx():
    with open('每日选股'+str(int(time.strftime("%Y%m%d")))+'.EBK') as f:
        d=f.read()
        wx_msg('wx87780fb826353ecc','XlOvsTpuNaINPsV2Wm6TMRLt_k1lgxZjJwrSaVND0Lo',1000003,d)
