from outlook.scheduler import OutlookScheduler
from outlook.storage import MongoToExcel

import time

spider = OutlookScheduler()
spider.run()
print('爬取完成！！！')

# time.sleep(10)
# save = MongoToExcel()
# save.run()

# print('存储完成！！！')

# time.sleep(6)

