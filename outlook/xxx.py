                # time.sleep(100)
                continue
                buyer_li_str = a[0].xpath('./table[1]/tbody[1]/tr[1]/td[1]/text()')
                buyer_li = [re.sub(r'\n', '', i) for i in buyer_li_str][1:-1]
                buyer = ''.join(buyer_li)

                note_li_str = a[0].xpath('./table[1]/tbody[1]/tr[1]/td[2]/text()')
                note = [re.sub(r'\n', '', i) for i in note_li_str][1]
                # print(note)
                address_li_str = a[0].xpath('./table/tbody[1]/tr[2]/td[1]/text()')
                address_li = [re.sub(r'\n', '', i)
                              for i in address_li_str][1:-3]
                address = ''.join(address_li)
                # print(address)
                url_li_str = a[0].xpath('./table[2]/tbody/tr/td[1]/a/@href')
                item_li = [re.search(r'item=(\d+)', urllib.parse.unquote(i)).group(1) for i in url_li_str]
                print(item_li)
                continue

                price_li_str = a[0].xpath('./table[2]/tbody/tr/td[2]/text()')
                price_li = [re.sub(r'\n|\s', '', i) for i in price_li_str][1:]
                # price_li = [float(i) for i in price_li]
                # print(price_li)
                qty_li_str = a[0].xpath('./table[2]/tbody/tr/td[3]/text()')
                qty_li = [re.sub(r'\n|\s', '', i) for i in qty_li_str][1:]
                qty_li = [int(i) for i in qty_li]
                # print(qty_li)
                amount_li_str = a[0].xpath('./table[2]/tbody/tr/td[4]/text()')
                amount_li = [re.sub(r'\n|\s', '', i)
                             for i in amount_li_str][1:]
                # print(amount_li)
                end_li_str = a[0].xpath('./table[3]/tbody/tr/td[2]/table/tbody/tr/td/text()')
                end_li = [re.sub(r'\n|\$|USD', '', i) for i in end_li_str]
                # print(end_li)
