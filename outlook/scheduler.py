#coding:utf-8
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
import outlook
import threading
import time
from lxml import etree
import re
import urllib.parse
from ebay import EbayApi
from ems import EMS
import json
import pymongo

class OutlookScheduler(object):
    def __init__(self):
        self.client = pymongo.MongoClient(host='127.0.0.1',port=27017)
        self.db = self.client['outlook']
        self.collection = self.db['data']

        self._outlook_email_queue = queue.Queue(10)
        self._content_email_queue = queue.Queue(10)
        self._collect_rsult_queue = queue.Queue(10)
        

    def decode_str(self,s):
        value, charset = decode_header(s)[0]
        if charset:
            value = value.decode(charset)
        return value

    def guess_charset(self,msg):
        charset = msg.get_charset()
        if charset is None:
            content_type = msg.get('Content-Type', '').lower()
            pos = content_type.find('charset=')
            if pos >= 0:
                charset = content_type[pos + 8:].strip()
        # print(charset)
        return charset

    # indent用于缩进显示:
    def print_info(self,msg, indent=0):
        if indent == 0:
            for header in ['From', 'To', 'Subject']:
                value = msg.get(header, '')
                if value:
                    if header=='Subject':
                        value = self.decode_str(value)
                    else:
                        hdr, addr = parseaddr(value)
                        # print(hdr,addr)
                        name = self.decode_str(hdr)
                        value = '%s <%s>' % (name, addr)
                # print('%s%s: %s' % ('  ' * indent, header, value))
        if (msg.is_multipart()):
            parts = msg.get_payload()
            for n, part in enumerate(parts):
                # print('%spart %s' % ('  ' * indent, n))
                # print('%s--------------------' % ('  ' * indent))
                self.print_info(part, indent + 1)
        else:
            content_type = msg.get_content_type()
            if content_type=='text/plain' or content_type=='text/html':
                try:
                    content = msg.get_payload(decode=True)
                    charset = self.guess_charset(msg)
                    if charset:
                        content = content.decode(charset)
                except Exception as e:
                    print(e)
                else:
                    # print(content)
                    self._content_email_queue.put(content)
            else:
                pass
                # print('%sAttachment: %s' % ('  ' * indent, content_type))

    def run_outlook(self):
        mail = outlook.Outlook()
        try:
            with open('usr.txt','r') as f:
                read_str = f.read()
                usr_dic = json.loads(read_str)
                usr = usr_dic['usr']
                pwd = usr_dic['pwd']
        except:
            usr = input('请输入您的outlook邮箱账号：')
            pwd = input('请输入您的密码：')
            usr_dic = json.dumps({'usr':usr,'pwd':pwd})
            with open('usr.txt','w') as f:
                f.write(usr_dic)
        finally:
            mail.login(usr,pwd)
            # mail.login('iwechen123@outlook.com','cw123456')
            mail.inbox()
            id_li = mail.allIds()
            print('本次筛选共有 %d 封有效邮件'%len(id_li))
            for index,ids in enumerate(id_li):
                msg_content = mail.getEmail(str(index+1))
                # print('----------------')
                # print(msg_content)
                self._outlook_email_queue.put(msg_content)

    def charset(self):
        while True:
            msg_content = self._outlook_email_queue.get()
            self.print_info(msg_content)

    def content_collec(self):
        while True:
            content = self._content_email_queue.get()
            html = etree.HTML(content)
            
            a = html.xpath('//*[@id="divtagdefaultwrapper"]/div/div[2]/table/tbody/tr/td[2]/table[1]/tbody/tr/td/table/tbody/tr[3]/td/table[2]/tbody/tr/td[2]/div/div/table/tbody/tr/td')
            if a == []:
                continue
            else:
                time = html.xpath('//*[@id="original-content"]/div[1]/div[2]/div[2]/text()')[0]
                # print(time)
                buyer_li_str = a[0].xpath('./table[1]/tbody[1]/tr[1]/td[1]/text()')
                buyer_li = [re.sub(r'\n','',i) for i in buyer_li_str][1:-1]
                buyer = ''.join(buyer_li)

                note_li_str = a[0].xpath('./table[1]/tbody[1]/tr[1]/td[2]/text()')
                note = [re.sub(r'\n','',i) for i in note_li_str][1]
                # print(note)
                address_li_str = a[0].xpath('./table/tbody[1]/tr[2]/td[1]/text()')
                address_li = [re.sub(r'\n','',i) for i in address_li_str][1:-3]
                address = ''.join(address_li)
                # print(address)
                url_li_str = a[0].xpath('./table[2]/tbody/tr/td[1]/a/@href')
                item_li = [re.search(r'item=(\d+)',urllib.parse.unquote(i)).group(1) for i in url_li_str]
                # print(url)
                price_li_str = a[0].xpath('./table[2]/tbody/tr/td[2]/text()')
                price_li = [re.sub(r'\n|\s','',i) for i in price_li_str][1:]
                # price_li = [float(i) for i in price_li]
                # print(price_li)
                qty_li_str = a[0].xpath('./table[2]/tbody/tr/td[3]/text()')
                qty_li = [re.sub(r'\n|\s','',i) for i in qty_li_str][1:]
                qty_li = [int(i) for i in qty_li]
                # print(qty_li)
                amount_li_str = a[0].xpath('./table[2]/tbody/tr/td[4]/text()')
                amount_li = [re.sub(r'\n|\s','',i) for i in amount_li_str][1:]
                # print(amount_li)
                end_li_str = a[0].xpath('./table[3]/tbody/tr/td[2]/table/tbody/tr/td/text()')
                end_li = [re.sub(r'\n|\$|USD','',i) for i in end_li_str]
                # print(end_li)
                weight_li = list()
                pid_li = list()
                
                # 理论总和
                ems_price_sum = 0
                air_price_sum = 0
                # print(item_li)
                googs_data_li = list()

                image_li = list()

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

                for index,pid in enumerate(item_li):
                    data_dict = dict()
                    data_dict['订单日期'] = time
                    data_dict['客户姓名'] = buyer
                    data_dict['客户邮箱'] = end_li[-1].split(' ')[-1]
                    data_dict['产品编号'] = googs_data_li[0+index]['pid']
                    data_dict['产品重量'] = googs_data_li[0+index]['weight']+'g'
                    data_dict['产品金额'] = price_li[0+index]
                    data_dict['订购数量'] = qty_li[0+index]
                    data_dict['应付总额'] = end_li[-4]
                    data_dict['实付总额'] = end_li[-2]
                    data_dict['实付运费'] = end_li[-8]
                    # 邮寄方式
                    if (float(end_li[-8])>29.2) or (weight_all > 700):
                         ems_method= '快递'
                    else:
                        if float(end_li[-8]) == 21.6:
                            ems_method= '航空'
                        else:
                            if sum(qty_li)==1:
                                ems_method= '快递'
                            else:
                                ems_method = '航空'

                    data_dict['邮寄方式'] = ems_method
                    # 返邮费
                    if (len(item_li))==1:
                        pay_back_price= 0
                    else:
                        # 1.判断大小
                        # 大大
                        if max(weight_li) > 700 and min(weight_li) > 700:
                            pay_back = 1 * 14  + 0 * 6
                        # 大小
                        elif max(weight_li) > 700 and min(weight_li) <= 700:
                            pay_back = 1 * 14  + 1 * 6
                        # 小小
                        elif max(weight_li) <= 700 and min(weight_li) <= 700:
                            pay_back = 0 * 14  + 1 * 6
                        # 2.判断邮寄方式
                        if ems_method == '航空':
                            # print(air_price_sum)
                            pay_postage = air_price_sum - pay_back
                            # print(pay_postage)
                        else:
                            pay_postage = ems_price_sum - pay_back
                        # 3，计算返还邮费金额
                        pay_back_price =round(float(end_li[-8]) -  pay_postage)
                        print(pay_back_price)

                    data_dict['返邮费金额'] = pay_back_price

                    data_dict['邮寄地址'] = address
                    data_dict['客户留言'] = note
                    data_dict['pid_li'] = pid_li
                    data_dict['image'] = image_li[0+index]

                    self._collect_rsult_queue.put(data_dict)
                    # print(data_dict)

    def save_to_mongo(self):
        while True:
            data = self._collect_rsult_queue.get()
            # print(data)
            try:
                self.collection.insert(data)
                print('mongo_sucessful')
            except:
                print('default')

    def init(self):
        t1 = threading.Thread(target=self.charset)
        # t1.setDaemon(True)
        t1.start()

        t2 = threading.Thread(target=self.content_collec)
        # t2.setDaemon(True)
        t2.start()

        t3 = threading.Thread(target=self.save_to_mongo)
        # t2.setDaemon(True)
        t3.start()

    def run(self):
        self.init()
        self.run_outlook()

    def main(self):
        self.run()

if __name__=="__main__":
    scheduler = OutlookScheduler()
    scheduler.main()
