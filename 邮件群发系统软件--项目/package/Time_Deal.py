import time

def getTimeNow():
    timeArray = time.localtime(time.time())
    TimeNow = time.strftime("%Y-%m-%d %H:%M:%S",timeArray)
    return TimeNow


