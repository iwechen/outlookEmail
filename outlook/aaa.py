import time,datetime

timeStr = '2018-05-10 12:24'

ltime = time.localtime(time.mktime(time.strptime(timeStr, "%Y-%m-%d %H:%M")))
timeStr = time.strftime("%Y-%m-%d %H:%M", ltime)
date_time = datetime.datetime.strptime(timeStr, '%Y-%m-%d %H:%M')


