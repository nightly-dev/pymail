#!/usr/bin/python
# -*- coding: utf-8 -*-

def write_log(msg):
    try:
        with open('mylog.html', 'a+') as f:
            f.write((msg+'<br>').encode('utf-8'))
    except:
        pass

def clear_log():
    try:
        with open('mylog.html', 'w+') as f:
            str = '<meta http-equiv=content-type content=text/html; charset=UTF-8 />'
            f.write(str)
    except:
        pass