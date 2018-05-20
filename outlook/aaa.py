import time,datetime

timeStr = 'Wednesday, May 2, 2018 2:06 AM'

ltime = time.localtime(time.mktime(time.strptime(timeStr, "%A, %b %d, %Y %H:%M %p")))
timeStr = time.strftime("%Y-%m-%d %H:%M", ltime)
date_time = datetime.datetime.strptime(timeStr, '%Y-%m-%d %H:%M')


print(timeStr)

