#-*- coding : utf-8 -*-
import requests,json
from . import Time_Deal

def getcurrentip():
    burl = 'https://jsonip.com'
    try:
        internet_ip = requests.get(burl,timeout=60).text
        return json.loads(internet_ip)['ip']
    except Exception as e:
        print("Warning：当前网络异常，网速太差或者断网。 @" + Time_Deal.getTimeNow())
        return ''