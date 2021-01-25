# -*- coding: utf-8 -*-
"""
Created on Wed Dec 23 19:35:57 2020

@author: 0100904
"""
#!/usr/bin/python

 

# from win32com.client import Dispatch

 

# xlApp = Dispatch('Excel.Application')

# xlApp.Visible = True

# xlbook = xlApp.Workbooks.open()
# xllbook.Add()

# xlApp.Worksheets.Add().Name = 'test'

# xlSheet = xlApp.Worksheets('test')

# xlSheet.Cells(1,1).Value = 'title'

# xlSheet.Cells(2,1).Value = 123
# xlBook.SaveAs(pwd + '\\demo.xlsx')
# xlbook.Close(True)

# xlApp.Quit()



#!/usr/bin/python
from win32com.client import Dispatch
import os
pwd = os.getcwd()
xlApp = Dispatch('Excel.Application')
xlApp.Visible = True
xlBook = xlApp.Workbooks.Add()
xlApp.Worksheets.Add().Name = 'test'
xlSheet = xlApp.Worksheets('test')
xlSheet.Cells(1,1).Value = 'title'
xlSheet.Cells(2,1).Value = 123
xlBook.SaveAs(pwd + '\\demo1.xlsx')
xlApp.Quit() # exit app