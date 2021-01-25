# -*- coding: utf-8 -*-
"""
Created on Tue Dec 22 20:14:01 2020

@author: 0100904
"""
my_id = 'yofc'
sql = ('''SELECT * FROM tmp_tabl where tmp_type in ('a','b','c') and id = '@id'
第二行xxx   and id = '@id'
第三行xxx and id ='@id'
#用replace批量替换变量
''').replace('@id', my_id) 
print(sql)