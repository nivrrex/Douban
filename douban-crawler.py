import requests
from pyquery import PyQuery as pyq
from lxml import etree
import openpyxl
import time
import re
import io, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer,encoding='utf8')

url = "http://douban.com/tag/股票?start=20&type=T"

def get_page_book_info(url,book):
    html = pyq(url)
    print('reading ...  {0}'.format(url))
    sys.stdout.flush()

    #获取url中的tag名称
    tag_split  = url.split("/")
    if tag_split != None:
        tag = tag_split[4]
        if str.find(tag,"?"):
            tag = tag.split("?")[0]
            print(tag)

    #获取页面中图书信息
    for element in html('ul.subject-list li.subject-item'):
        link = pyq(element)('div.info h2 a').attr('href')
        title = pyq(element)('div.info h2 a').attr('title')
        title_info = pyq(element)('div.info h2 a').text()
        pic = pyq(element)('div.pic a.nbg img').attr('src')
        pub = pyq(element)('div.info div.pub').text()
        rating_nums = pyq(element)('div.info div.star span.rating_nums').text()
        pinglun = pyq(element)('div.info div.star span.pl').text()
        info = pyq(element)('div.info p').text()
        buy_info = pyq(element)('div.cart-actions span.buy-info a').text()
        
        index = len(book) + 1
        book[index] = {}
        book[index]["Tag"] =  tag
        book[index]["Index"] =  index
        book[index]["Title"] =  title
        book[index]["Title Info"] =  title_info
        book[index]["Link"] =  link
        book[index]["Picture"] =  pic
        book[index]["PUB"] =  pub
        book[index]["Rating"] =  rating_nums
        book[index]["Pinglun"] =  pinglun
        book[index]["Info"] =  info
        book[index]["Buy Info"] =  buy_info

    #获取页面中下一页链接
    next_link = html('div.paginator span.next a').attr("href")
    if next_link != None:
        next_link = "https://book.douban.com" + next_link
    else:
        return book, None
        #print(book[index])
        #print("\n")
    return book, next_link

def get_douban_movies_now():
    url = 'https://movie.douban.com/'
    html = pyq(url)
    print('reading ...  https://movie.douban.com/\n')
    sys.stdout.flush()

    movies = {}
    for element in html('div.s div.screening-bd ul.ui-slide-content li.ui-slide-item'):
        #获取电影Title
        title = pyq(element)('ul li.title a').text()
        if title:
            movies[title] = {}
            movies[title]["Title"] = title
        #获取电影评价
        rating = pyq(element)('ul li.rating span').text()
        if rating:
            movies[title]["Rating"] = rating
    return movies
    



if __name__ == "__main__":
    #获取最新的豆瓣电影评价
    #movies = get_douban_movies_now()
    #for v in movies :
    #    print(movies[v]["Title"],movies[v]["Rating"])

    #获取所有豆瓣的所有tag下的书籍的相关资料
    url = "https://book.douban.com/tag/"
    html = pyq(url)

    #book的hash数据写入文件，写入标题头
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Index","Tag","Title","Title Info","Link","Picture","PUB","Rating","Pinglun","Info","Buy Info"])

    #根据Tag页面信息获取所有tag,然后逐一抓取
    for element in html('div.grid-16-8 div.article div div table tbody tr td a'):
        tag = pyq(element).text()
        tag_link = pyq(element).attr('href')
        url = "https://book.douban.com" + tag_link
        print(tag,tag_link,url)

        book = {}
        while True:
            #获取页面中book信息
            book,next_link = get_page_book_info(url,book)
            if next_link != None:
                url = next_link
            else:
                break
            time.sleep(0.8)

        #写入文件
        for index in book:
            ws.append([book[index]["Index"], book[index]["Tag"], book[index]["Title"], book[index]["Title Info"], book[index]["Link"], \
            book[index]["Picture"], book[index]["PUB"], book[index]["Rating"], book[index]["Pinglun"], book[index]["Info"], book[index]["Buy Info"]])
        #保存文件
        wb.save("book_store_douban.xlsx")

        print("\n")
        time.sleep(1.2)
