# coding:utf-8
'''
Created on 2018年5月10日
@author: chenwei
Email:iwechen123@gmail.com
'''

import imaplib
from email.header import decode_header
from email.utils import parseaddr
import poplib
from email.parser import Parser
from six.moves import queue
from .outlook import Outlook
import threading
import time
import datetime
from lxml import etree
import re
import urllib.parse
from .ebay import EbayApi
from .ems import EMS
import json
import pymongo
import hashlib

class OutlookScheduler(object):
    def __init__(self):
        self.flag = True
        self.client = pymongo.MongoClient(host='127.0.0.1',port=27017)
        self.db = self.client['outlook']
        self.collection = self.db['data']

        self._outlook_email_queue = queue.Queue(10)
        self._content_email_queue = queue.Queue(10)
        self._collect_rsult_queue = queue.Queue(10)
        self._order_data_dic_queue = queue.Queue(50)

    def decode_str(self, s):
        value, charset = decode_header(s)[0]
        if charset:
            value = value.decode(charset)
        return value

    def guess_charset(self, msg):
        charset = msg.get_charset()
        if charset is None:
            content_type = msg.get('Content-Type', '').lower()
            pos = content_type.find('charset=')
            if pos >= 0:
                charset = content_type[pos + 8:].strip()
        # print(charset)
        return charset

    # indent用于缩进显示:
    def print_info(self, msg, indent=1):
        if (msg.is_multipart()):
            parts = msg.get_payload()
            for n, part in enumerate(parts):
                self.print_info(part, indent + 1)
        else:
            content_type = msg.get_content_type()
            if content_type == 'text/plain' or content_type == 'text/html':
                try:
                    content = msg.get_payload(decode=True)
                    charset = self.guess_charset(msg)
                    if charset:
                        content = content.decode(charset)
                except Exception as e:
                    print(e)
                else:
                    self._content_email_queue.put(content)

    def run_outlook(self):
        mail = Outlook()
        try:
            with open('usr.txt', 'r') as f:
                read_str = f.read()
                usr_dic = json.loads(read_str)
                usr = usr_dic['usr']
                pwd = usr_dic['pwd']
        except:
            usr = input('请输入您的outlook邮箱账号：')
            pwd = input('请输入您的密码：')
            usr_dic = json.dumps({'usr': usr, 'pwd': pwd})
            with open('usr.txt', 'w') as f:
                f.write(usr_dic)
        finally:
            mail.login(usr, pwd)
            # mail.login('iwechen123@outlook.com','cw123456')
            mail.inbox()
            id_li = mail.allIds()
            # print(id_li)
            print('本次筛选共有 %d 封有效邮件' % len(id_li))
            for index, ids in enumerate(id_li):
                # print(ids)
                msg_content = mail.getEmail(str(index+1))
                self._outlook_email_queue.put(msg_content)

    def charset(self):
        while self.flag:
            msg_content = self._outlook_email_queue.get()
            self.print_info(msg_content)

    def time_str_to_dtime(self,timeStr):
        # print(timeStr)  May 3, 2018 19:46
        ltime = time.localtime(time.mktime(time.strptime(timeStr, "%A, %b %d, %Y %H:%M")))
        timeStr = time.strftime("%Y-%m-%d %H:%M", ltime)
        date_time = datetime.datetime.strptime(timeStr, '%Y-%m-%d %H:%M')
        return {'time_str':timeStr,'date_time':date_time}

    def hash_to_md5(self,sign_str):
        m= hashlib.md5()
        sign_str = sign_str.encode('utf-8')
        # 加密字符串  
        m.update(sign_str) 
        sign = m.hexdigest() 
        return sign

    def submit(self,weight_all):
        if weight_all <= 50:
            return 10
        elif 51 <= weight_all <= 100:
            return 20
        elif 101 <= weight_all:
            return 30

    def ems_method(self,pay_ems,weight_all,qty_li):
        # 邮寄方式
        if (pay_ems > 29.2) or (weight_all > 700):
            ems_method = '快递'
        else:
            if pay_ems == 21.6:
                ems_method = '航空'
            else:
                if sum(qty_li) == 1:
                    ems_method = '快递'
                else:
                    ems_method = '航空'
        return ems_method

    def pay_back_price(self,ems,pay_ems,item_li,weight_li,ems_method,qty_li,air_price_sum,ems_price_sum):
        if (len(item_li)) == 1:
            pay_back_price = 0
        else:
            # 1.判断大小
            # 大大   (总数 - 1) * 14
            if max(weight_li) > 700 and min(weight_li) > 700:
                pay_back = (sum(qty_li)-1) * 14 + 0 * 6
                # 判断邮寄方式
                if ems_method == '航空':
                    # 理论总和 - 返费金额
                    pay_postage = air_price_sum - pay_back
                else:
                    # 理论总和 - 返费金额
                    pay_postage = ems_price_sum - pay_back

            # 大小  (大 - 1) * 14 + 小 * 6
            elif max(weight_li) > 700 and min(weight_li) <= 700:
                # 大 重量列表
                max_li = list()
                # 小 重量列表
                min_li = list()
                for weight in weight_li:
                    if weight > 700:
                        max_li.append(weight)
                    else:
                        min_li.append(weight)
                pay_back = (len(max_li) - 1) * 14 + len(min_li) * 6
                # 定义混合 理论总和
                mix_price_sum =0
                # 大 默认快递 理论总和
                ems_price_dic = ems.functions(sum(max_li))
                mix_price_sum += ems_price_dic['ems']
                # 小 默认航空 理论总和
                ems_price_dic = ems.functions(sum(min_li))
                mix_price_sum += ems_price_dic['air']
                # 理论总和 - 返费金额
                pay_postage = mix_price_sum - pay_back

            # 小小  (总数 - 1) * 6
            elif max(weight_li) <= 700 and min(weight_li) <= 700:
                pay_back = 0 * 14 + (sum(qty_li)-1) * 6
                # 2.判断邮寄方式
                if ems_method == '航空':
                    # 理论总和 - 返费金额
                    pay_postage = air_price_sum - pay_back
                else:
                    # 理论总和 - 返费金额
                    pay_postage = ems_price_sum - pay_back
            # 3，计算返还邮费金额
            pay_back_price = round(pay_ems - pay_postage)
        return pay_back_price

    def content_collec(self):
        error_name = 0
        while self.flag:
            content1 = self._content_email_queue.get()
            content2 = re.sub(r'\r|\n|\<.*?\>|\&nbsp\;|\s{3,}','',content1)
            content = re.sub(r'\<.*?\>','',content2,re.S)
            try:
                order_li_one = re.findall(r'Item\#\s+(\d+).*?(\$\d+\.\d+ USD).*?(\d+).*?(\$\d+\.\d+ USD)',content)
            except Exception as e:
                print('提取第一种模板出错',e)
            if order_li_one == []:
                error_name += 1
                file_name = 'error_%d'%error_name
                with open(file_name+'.html', 'w') as f:
                    f.write(content1)
                continue
            else:
                # 数据保存字典
                order_data_dic = dict()
                # 订单ID
                sign_str = json.dumps(order_li_one) + str(int((time.time()*1000)))
                sign = self.hash_to_md5(sign_str)
                # 订单日期   May 2, 2018 2:06
                time_dic = self.time_str_to_dtime(re.findall(r'Sent:\s+(.*?)\s+[A,P,M]{2}',content)[0])
                # 客户姓名
                buyer_str = re.sub(r'<.*?>|\r|\n','',re.findall(r'Buyer(.*?)Note',content,re.S)[0])
                # 客户邮箱
                buyer_email = re.findall(r'From:\s{0,1}(.*?\.com)',content)[0]
                # 产品编号
                item_li =list()
                # 产品金额
                price_li = list()
                # 订购数量
                qty_li = list()
                for order in order_li_one:
                    # 添加产品编号列表
                    item_li.append(order[0])
                    # 添加单价列表
                    price_li.append(order[1])
                    # 添加购买数量列表
                    qty_li.append(int(order[2]))
                # 实付运费
                pay_ems = float(re.findall(r'Shipping and handling\$(\d+\.\d+)\s+USD',content)[0])
                # 应付总额
                total = float(re.findall(r'Total\s{0,2}\$(\d+\.\d+)\s+USD',content)[0])
                # 实付总额
                payment = float(re.findall(r'Payment\s{0,2}\$(\d+\.\d+)\s+USD',content)[0])
                # 邮寄地址
                address = re.sub(r'\r|\n','',re.findall(r'Shipping address(.*?)Shipping details',content,re.S)[0])
                # 客户留言
                note = re.sub(r'\r|\n','',re.findall(r'Note to seller(.*?)Shipping address',content,re.S)[0])
                
                order_data_dic['sign'] = sign
                order_data_dic['time_dic'] = time_dic
                order_data_dic['buyer_str'] = buyer_str
                order_data_dic['buyer_email'] = buyer_email
                order_data_dic['item_li'] = item_li
                order_data_dic['price_li'] = price_li
                order_data_dic['qty_li'] = qty_li
                order_data_dic['pay_ems'] = pay_ems
                order_data_dic['total'] = total
                order_data_dic['payment'] = payment
                order_data_dic['address'] = address
                order_data_dic['note'] = note
                # print(item_li)
                # 添加到任务队列
                self._order_data_dic_queue.put(order_data_dic)

    def count_data(self):
        while True:
            order_data_dic = self._order_data_dic_queue.get()

            weight_li = list()
            pid_li = list()
            # 理论总和
            ems_price_sum = 0
            air_price_sum = 0
            googs_data_li = list()

            image_li = list()
            item_li = order_data_dic['item_li']
            for item in item_li:
                ebay = EbayApi()
                googs_data = ebay.load_api(item)
                # print(googs_data)
                googs_data_li.append(googs_data)

                weight_li.append(float(googs_data['weight']))
                pid_li.append(googs_data['pid'])
                image_li.append(googs_data['image'])
                # 理论总和
                ems = EMS()
                ems_price_dic = ems.functions(float(googs_data['weight']))
                # print(ems_price_dic['air'])
                ems_price_sum += ems_price_dic['ems']
                air_price_sum += ems_price_dic['air']
            # 总重量
            weight_all = sum(weight_li)
            qty_li = order_data_dic['qty_li']
            for index, item in enumerate(item_li):
                data_dict = dict()
                data_dict['sign'] = order_data_dic['sign']
                data_dict['订单日期'] = order_data_dic['time_dic']['time_str']
                data_dict['客户姓名'] = order_data_dic['buyer_str']
                data_dict['客户邮箱'] = order_data_dic['buyer_email']
                data_dict['产品编号'] = googs_data_li[0+index]['pid']
                data_dict['产品重量'] = googs_data_li[0+index]['weight']+'g'
                data_dict['产品金额'] = order_data_dic['price_li'][0+index]
                data_dict['订购数量'] = order_data_dic['qty_li'][0+index]
                data_dict['应付总额'] = order_data_dic['total']
                data_dict['实付总额'] = order_data_dic['payment']
                pay_ems = order_data_dic['pay_ems']
                data_dict['实付运费'] = pay_ems
                data_dict['邮寄地址'] = order_data_dic['address']
                data_dict['客户留言'] = order_data_dic['note']
                qty_li = order_data_dic['qty_li']
                # 邮寄方式
                ems_method = self.ems_method(pay_ems,weight_all,qty_li)
                data_dict['邮寄方式'] = ems_method
                # 返邮费
                pay_back_price = self.pay_back_price(ems,pay_ems,item_li,weight_li,ems_method,qty_li,air_price_sum,ems_price_sum)
                data_dict['返邮费金额'] = pay_back_price

                data_dict['产品编号集合'] = ','.join(pid_li)
                weight_li_str = [str(i)+'g' for i in weight_li]

                data_dict['产品重量集合'] = ','.join(weight_li_str)
                data_dict['订购总数'] = sum(order_data_dic['qty_li'])
                data_dict['image'] = image_li[0+index]
                data_dict['date_time'] = order_data_dic['time_dic']['date_time']
                data_dict['报'] = self.submit(weight_all)

                self._collect_rsult_queue.put(data_dict)
                # print(data_dict)

    def save_to_mongo(self):
        while self.flag:
            data_dict = self._collect_rsult_queue.get()
            try:
                self.collection.insert(data_dict)
                print('save_to_mongo_successful!!!')
            except:
                print('default')


    def stop_threading(self):
        self.flag = True
        count = 0
        # while True:
        #     # print(count)
        #     if count < 30:
        #         if self._collect_rsult_queue.empty() and self._content_email_queue.empty() and self._outlook_email_queue.empty():
        #             count += 1
        #             time.sleep(2)
        #         else:
        #             time.sleep(1)
        #             count = 0
        #     else:
        #         print('stop_byby!!!')
        #         self.flag = False
        #         break

    def init(self):
        t0 = threading.Thread(target=self.stop_threading)
        # t0.setDaemon(True)
        t0.start()

        t1 = threading.Thread(target=self.charset)
        # t1.setDaemon(True)
        t1.start()

        t2 = threading.Thread(target=self.content_collec)
        # t2.setDaemon(True)
        t2.start()

        t3 = threading.Thread(target=self.save_to_mongo)
        # t3.setDaemon(True)
        t3.start()

        t4 = threading.Thread(target=self.count_data)
        # t4.setDaemon(True)
        t4.start()

    def run(self):
        self.init()
        # self.stop_threading()
        self.run_outlook()

    def main(self):
        self.run()


if __name__ == "__main__":
    scheduler = OutlookScheduler()
    scheduler.main()
