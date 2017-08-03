#coding:utf-8
#this is based on Python3.6

import json

file_desc = open('rules.json', 'r', encoding='UTF-8')
dict_str = json.load(file_desc)
file_desc.close()

print(dict_str)
print(len(dict_str))