import requests
from lxml import etree
import re

class EbayApi(object):
    def __init__(self):
        self.headers = {
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1',
            'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/66.0.3359.139 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
            'Referer': 'https://www.ebay.com/itm/292544733016?ViewItem=&item=292544733016&ppid=PPX000600&cnac=C2&rsta=en_C2(en_C2)&cust=1GT93470EG729545P&unptid=546999de-4dd3-11e8-81f2-441ea14e5834&t=&cal=f9c0ca85adcd9&calc=f9c0ca85adcd9&calf=f9c0ca85adcd9&unp_tpcid=email-auction-payment-notification&page=main:email&pgrp=main:email&e=op&mchn=em&s=ci&mail=sys&data=02%7C01%7C%7Cf492119e61b740b5caf908d5aff73d1a%7C84df9e7fe9f640afb435aaaaaaaaaaaa%7C1%7C0%7C636608398769505177&sdata=VKf/VQVIIRGnUPuEokYte9HifctVwKpIRyqBC28WRFk=&reserved=0',
            'Accept-Encoding': 'gzip, deflate, br',
            'Accept-Language': 'zh,zh-CN;q=0.9',
        }

    def load_api(self,item):
        params = {
            'ViewItemDescV4': '',
            'item': item,
            't': '0',
            'tid': '10',
            'category': '19268',
            'seller': 'rikoo*',
            'excSoj': '1',
            'excTrk': '1',
            'lsite': '0',
            'ittenable': 'false',
            'domain': 'ebay.com',
            'descgauge': '1',
            'cspheader': '1',
            'oneClk': '1',
            'secureDesc': '1',
        }
        url = 'https://vi.vipr.ebaydesc.com/ws/eBayISAPI.dll'

        response = requests.get(url=url, headers=self.headers, params=params).content.decode('utf-8')
        html = etree.HTML(response)
        return_dict = dict()
        try:
            weight_str = html.xpath('//*[@id="ds_div"]/div[1]/div/div[2]/div[1]/div[2]/dl/dd[2]/text()')[0]
            image = html.xpath('//*[@id="ds_div"]/div[1]/div/div[2]/div[1]/div[2]/div[2]/p[1]/img/@src')[0]
            weight = re.search(r'(\d+)\s{0,}g',weight_str).group(1)
            pid = re.search(r'_(\d+)\.',image).group(1)
        except Exception as e:
            print('商品重量信息出错！！！',e)
            image = item
            weight = '0'
            pid = item
        finally:
            return_dict['pid'] = pid
            return_dict['weight'] = weight
            return_dict['image'] = image
            return return_dict

    def run(self):
        item = '292544733016'
        self.load_api(item)

    def main(self):
        self.run()

if __name__=='__main__':
    ebay = EbayApi()
    ebay.main()




