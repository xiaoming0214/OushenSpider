#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import string

from _datetime import datetime
import urllib3
import requests
import xlwt
from bs4 import BeautifulSoup

__author__ = 'zhengxiaoming'
__date__ = '2018/11/9'

url = 'https://mp.weixin.qq.com/s/tTrYkEa21VMWPl9uhqftwQ'
itemResponse = []
keyword = None

def requsetIndex():
    urllib3.disable_warnings()
    r = requests.get(url, verify=False)
    r.encoding = 'utf-8'
    return r.text

def parseIndex(response):
    indexResult = []
    soup = BeautifulSoup(response,'html.parser')
    for element in soup.find_all('a'):
        link = element['href']
        if 'http' in link:
            data = {}
            data['href'] = element['href']
            data['name'] = element.text
            indexResult.append(data)

    return indexResult

def reuestAllItem(indexResult):
    for item in indexResult:
        requestItem(item['href'])

    print('End')


def requestItem(url):
    print("开始请求",url)
    urllib3.disable_warnings()
    r = requests.get(url, verify=False)
    r.encoding = 'utf-8'
    parseItem(r.text)

def parseItem(response):
    soup = BeautifulSoup(response,'html.parser')
    itemResult = []
    filterResult = []
    process = None
    for item in soup.find_all('span'):
        try:
            if 'style' in item.attrs:
                if 'color: rgb(122, 122, 122)' in item.attrs['style'] :
                    resultItem = {}
                    resultItem['question'] = item.contents[0]
                    process = True
                    itemResult.append(resultItem)
                if 'color: rgb(0, 176, 80)' in item.attrs['style'] :
                    if process :
                        resultItem = itemResult[-1]
                        resultItem['answer'] = item.contents[0]
                        process = None
        except:
            pass

    for item in itemResult :
        if 'answer' in item and 'question' in item :
            if item['question'] != None and item['answer'] != None:
                try:
                    if keyword in item['question'].upper() or  keyword in item['answer'].upper() :
                        filterResult.append(item)
                except Exception as e:
                    print(e)

    print(filterResult)
    itemResponse.extend(filterResult)



# 写入excel
def reportToExcel(itemResult):
    # 实例化一个Workbook()对象(即excel文件)
    wbk = xlwt.Workbook()
    # 新建一个名为Sheet1的excel sheet。此处的cell_overwrite_ok =True是为了能对同一个单元格重复操作。
    sheet = wbk.add_sheet('Sheet1', cell_overwrite_ok=True)
    # 遍历result中的没个元素。
    i = 0
    if len(itemResult) > 0:
        for item in itemResult:
            try:
                if 'answer' in item and 'question' in item:
                    sheet.write(i, 0, item['question'])
                    sheet.write(i, 1, item['answer'])
                    i = i + 1
            except:
                pass

        # 获取当前日期，得到一个datetime对象如：(2016, 8, 9, 23, 12, 23, 424000)
        today = datetime.today()
        # 将获取到的datetime对象仅取日期如：2016-8-9
        today_date = datetime.date(today)

        # 以传递的name+当前日期作为excel名称保存。
        wbk.save("oushen-" +keyword + "-"+ str(today_date) + '.xls')

if __name__ == '__main__':

    keyword = 'HZ'

    response = requsetIndex()
    indexResult = parseIndex(response)
    reuestAllItem(indexResult)
    print(itemResponse)
    reportToExcel(itemResponse)

    pass

