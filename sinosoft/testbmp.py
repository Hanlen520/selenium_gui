#!/usr/bin/python
#coding=UTF-8
import datetime
import time

def getday(nowdate,n=1):
    y = int(nowdate.split("-",2)[0])
    m = int(nowdate.split("-", 2)[1])
    d = int(nowdate.split("-", 2)[2])
    the_date = datetime.datetime(y,m,d)
    result_date = the_date + datetime.timedelta(days=n)
    d = result_date.strftime('%Y-%m-%d')
    return d
print(getday(time.strftime('%Y-%m-%d'),182)) #8月15日后21天
# print(getday(2017,9,1,-10)) #9月1日前10天