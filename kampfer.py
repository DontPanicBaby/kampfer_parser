#! /usr/bin/python
# -*- coding: utf-8 -*-
'''
One-time parser kampfer.ru
'''

from grab import Grab
import os, re
import lxml.html
from grab.error import GrabNetworkError
from itertools import count
import string
from datetime import datetime
from  xlwt import *
import sys

print '[START]', datetime.now()


if not os.path.exists('images/'):
    os.mkdir('images/')

g = Grab()

done = []
try:
    done_file = open('done', 'a+')
    for i in done_file.readlines():
        done.append(i.strip())
except:
    print "[ERROR] file was not opened, app has closed"
    exit(0)
    

g.go('http://kampfer.ru/')

categories = []

for i in g.tree.xpath('//div[@class="main-cat"]/table/tr/td/h2/a/@href'): 
    categories.append('http://kampfer.ru' + i)


wb = Workbook()
ws0 = wb.add_sheet('0')

id_count = count(1)
rownum = 1
while True:
    if not categories: break
    url = categories.pop()
    if url in done: continue
    done.append(url)
    g.go(url)
    next_pages = g.tree.xpath(u'//a[contains(text(),"след >>")]/@href')
    if next_pages and next_pages[0] not in categories and next_pages[0] not in done: 
        categories.insert(0, 'http://kampfer.ru'  + next_pages[0])
    for i in g.tree.xpath('//a[contains(@href, "/product/")]/@href'):
        product_url = 'http://kampfer.ru' + i
        if product_url in done: continue
        number = 1000 + id_count.next()
        try:
            g.go(product_url)
        except GrabNetworkError: 
            print '[ERROR] Fail loading page: %s' %product_url
            continue
        done_file.write(product_url + '\n')
        print '[DONE]', product_url
        doc = g.tree 
        
        title = ''.join(doc.xpath('//div[@class="cpt_product_name"]//h2/text()')).strip()
        if not title: title = ''.join(doc.xpath('//div[@class="cpt_product_name"]//h1/text()')).strip()
        if not title: title = ''.join(doc.xpath('//title/text()')).strip().encode('utf-8')
        try:
            price = float(''.join(doc.xpath('//div[@class="cpt_product_price"]/p[@class="rrc"]/strong/text()')).replace(' ', '').replace(u'руб', ''))
        except: price = '0'
        article = ''.join(doc.xpath(u'//b[text()="Артикул: "]/following-sibling::text()')).strip()
        desc_tag = doc.xpath(u'//div[@class="cpt_product_description"]/div')
        if desc_tag:
            desc = lxml.html.tostring(desc_tag[0],  encoding='utf-8').replace('<br>', '\n')
            desc = re.sub('<a>.+?</a>', '', desc)
            desc = re.sub('<[^>]*>', '', desc)
            desc = desc.decode('utf-8')
        else: desc = ''
        image =  'http://kampfer.ru' + ''.join(doc.xpath('//*[@id="img-current_picture"]/@src'))
        image_name = '%s.%s' %(number, image.split('.')[-1])
        category = ''.join(doc.xpath('//li[@class="child current" or @class=" current"]/a/text()')).strip()
        count_tabs = len(doc.xpath('//li[@class="child current" or @class=" current"]/a/img'))
        if count_tabs >= 2 and category[0].upper() in string.ascii_uppercase:
            try:
                category = doc.xpath('//li[@class="child current" or @class=" current"]/preceding-sibling::li/a[count(img)=%s]/text()' %(count_tabs-1))[-1].strip()
            except: 
                category = ''
                print '[ERROR] category'
                
        try: g.download(image, os.path.join('images',  image_name))
        except GrabNetworkError:
            print 'Fake download image'
        except IOError:
            print 'IOError'
        image_counter = count(1)
        image_number = image_counter.next()
        for extimageurl in doc.xpath('//div[@class="dopf"]//img/@src'):
            try: 
                g.download('http://kampfer.ru' + extimageurl, os.path.join('images',  '%s_%s.%s' %(number, image_number, extimageurl.split('.')[-1])))
                image_number = image_counter.next()
            except GrabNetworkError: continue
            except IOError: print 'IOError'
        ws0.write(rownum, 0, number)
        ws0.write(rownum, 1, number)
        ws0.write(rownum, 2, number)
        ws0.write(rownum, 3, title)
        ws0.write(rownum, 4, int(price))
        ws0.write(rownum, 8, desc)
        ws0.write(rownum, 9, category)
        ws0.write(rownum, 13, "new")
        ws0.write(rownum, 14, 0)
        ws0.write(rownum, 15, 0)
        rownum += 1
        done.append(product_url)
     
if len(sys.argv) == 2:  wb.save(sys.argv[0]) 
else: wb.save('export.xls')
done_file.close()
print '[DONE]', datetime.now()

