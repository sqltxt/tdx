import time
import queue
import requests
import json

def wx_msg(corp_id, secret,agentid,msg):
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

with open('每日选股'+str(int(time.strftime("%Y%m%d")))+'.EBK') as f:
    d=f.read()
    print (d)
    wx_msg('wx87780fb826353ecc','WVClMaoadG7yP6sbsEox5HCwuHFG1SjrDqIjMVVfQmc',1000002,d)

