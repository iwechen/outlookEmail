# coding:utf-8
'''
Created on 2018年5月15日
@author: chenwei
Email:iwechen123@gmail.com
'''
import requests
import pymongo
import sys
import os
import csv
import xlwt
from docx import Document
from docx.shared import Pt
from docx.shared import Inches
from docx.oxml.ns import qn
import json
import re


class MongoToExcel(object):
    def __init__(self):
        self.client = pymongo.MongoClient(host='127.0.0.1', port=27017)
        self.db = self.client['outlook']
        self.collection1 = self.db['data']

        self.iamge_list = list()

    def save_to_excel(self, data_li):
        path = sys.path[0]+'/2018-5-10/'
        folder = os.path.exists(path)
        if not folder:
            os.makedirs(path)
        # 创建一个Workbook对象，这就相当于创建了一个Excel文件
        book = xlwt.Workbook(encoding='utf-8', style_compression=0)

        sheet = book.add_sheet('test', cell_overwrite_ok=True)
        for index, data in enumerate(data_li):
            if index == 0:
                for i, k in enumerate(list(data.keys())):
                    sheet.write(index, i, k)
                for i, k in enumerate(list(data.values())):
                    sheet.write(index+1, i, k)
            else:
                for i, k in enumerate(list(data.values())):
                    sheet.write(index+1, i, k)

        book.save(path+'订单列表.xls')

    def save_to_word(self, data_list):
        # 打开文档
        sss = ''
        document = Document()
        for data in data_list:
            a0 = data['订单日期']
            a = ','.join(data['pid_li'])
            b = data['返邮费金额']
            c = data['客户姓名']
            d = data['邮寄地址']
            ww = '%s 订单信息\n产品号：%s     报：%d   \n 客户姓名：%s \n 邮寄地址：%s \n\n' % (
                a0, a, b, c, d)
            sss += ww

        paragraph = document.add_paragraph(sss)
        path = sys.path[0] + '/2018-5-10/'
        document.save(path+'订单列表信息.docx')

    def find_mongo(self):
        f = open('file1.csv', 'r')
        datas = f.read()
        f.close()

        data_str_li = re.sub(r'}{', '}&&&{', datas).split('&&&')
        data_li = list(set(data_str_li))
        data_list = list()
        for i in data_li:
            # print(i)
            data_dict = json.loads(i)
            data_list.append(data_dict)
            self.iamge_list.append(
                {"pid": data_dict['产品编号'], 'image': data_dict['image']})

        if data_list == []:
            return

        self.save_to_excel(data_list)
        print('Excel successful!!!')
        self.save_to_word(data_list)
        print('Word successful!!!')
        self.down_load_image()

    def down_load_image(self):
        for data_dic in self.iamge_list:
            ima_name = data_dic['pid']
            url = data_dic['image']
            print('%s down_load!!!' % ima_name)
            path = sys.path[0]+'/2018-5-10/2018-5-10/'
            folder = os.path.exists(path)
            if not folder:
                os.makedirs(path)
            response = requests.get(url).content
            with open(path + ima_name+'.jpg', 'wb') as f:
                f.write(response)

    def run(self):
        self.find_mongo()

    def main(self):
        self.run()


if __name__ == "__main__":
    excel = MongoToExcel()
    excel.main()
