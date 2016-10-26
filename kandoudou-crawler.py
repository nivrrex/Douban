#-*- coding:utf-8 -*
import requests
from pyquery import PyQuery as pyq
from lxml import etree
import openpyxl
import time
import io, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer,encoding='utf8')


def get_page_book_info(url,book):
    html = pyq(url)
    next_link = None

    print('reading ...  {0}\n'.format(url))
    sys.stdout.flush()
    
    #获取图书信息
    for element in html('ul.list li.o'):
        o_img = pyq(element)('div.o-img')        
        o_info = pyq(element)('div.o-info')
        
        link = o_img('a').attr('href')
        img_src = o_img('img').attr('src')
        o_name = pyq(element)('h3.o-name a').text()
        o_author = pyq(element)('p.o-author a').text()
        o_ext = pyq(element)('p.o-ext').text()
        o_cate = pyq(element)('p.o-cate a').text()
        o_data = pyq(element)('p.o-data i.icon').text()
        t_temp  = o_data.split(" ")
        if t_temp != None:
            o_click = t_temp[0]
            o_download = t_temp[1]
        print (o_name,o_author,link,img_src,o_ext,o_cate,o_click,o_download)
        sys.stdout.flush()

        index = len(book) + 1
        book[index] = {}
        book[index]["Index"] =  index
        book[index]["Name"] =  o_name
        book[index]["Author"] =  o_author
        book[index]["Tag"] =  o_cate        
        book[index]["EXT"] =  o_ext
        book[index]["Link"] =  link
        book[index]["Picture"] =  img_src
        book[index]["Click_Number"] =  o_click
        book[index]["Download_Number"] =  o_download

    #获取页面中下一页链接
    for link in html('ul.paging li a'):
        if pyq(link).text() == '下一页':
            next_link = pyq(link).attr('href')

    if next_link != None:
        return book, next_link
    else:
        return book, None


if __name__ == "__main__":
    #book的hash数据写入文件，写入标题头
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Index","Name","Tag","Author","EXT","Link","Picture","点击数量","下载数量"])
    
    
    #获取所有看看豆的下的书籍的相关资料
    url = "http://kankandou.com/"
    book = {}
    while True:
        #获取页面中book信息
        book,next_link = get_page_book_info(url,book)
        if next_link != None:
            url = next_link
        else:
            break
        print("\n")
        time.sleep(1.2)

    #写入文件
    print('Start to write file ... \n')
    sys.stdout.flush()

    for index in book:
        ws.append([book[index]["Index"], book[index]["Name"], book[index]["Tag"], book[index]["Author"], book[index]["EXT"], \
                    book[index]["Link"], book[index]["Picture"], book[index]["Click_Number"], book[index]["Download_Number"]])
    #保存文件
    wb.save("book_store_kandoudou.xlsx")

    print('End to write file ... \n')
    sys.stdout.flush()

