#!/usr/bin/python
# -*- coding: utf-8 -*-
import json
import os
import time

def load_cfg():
    try:
        with open('mycfg.json', 'r') as f:
            data = json.load(f)
            #print data
            return data
    except:
            return None

if __name__ == '__main__':
    load_cfg()
    time.sleep(6)