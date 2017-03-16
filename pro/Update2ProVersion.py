# -*- coding: utf-8 -*-
"""
Created on Thu Mar 16 22:59:02 2017

@author: Administrator
"""


import os
import re
import sqlite3
import openpyxl
import sys
reload(sys)
sys.setdefaultencoding("utf-8")

#os.chdir(sys.path[0])


db=sqlite3.connect("dictionary.db")
db.text_factory=str
cu=db.cursor()

#excelname="tolearn.xlsm"
excelname="mastered.xlsm"

wb=openpyxl.load_workbook(filename=excelname,keep_vba=True)


print u"processing..."


# tolearn.xslx -> mastered.xslx
for sheet in wb.get_sheet_names():
    ws=wb[sheet]
    row_count=0
    for row in ws.rows:
        row_count+=1
        if row_count>1:
            command="select en, us, en_voice, us_voice, example from english_dictionary where word = ? COLLATE NOCASE"            
            cu.execute(command,(row[0].value,))
            res=cu.fetchone()
            if res:
                for i in range(len(res)):
                    ws.cell(row=row_count,column=i+5).value=res[i]

wb.save(excelname)

print u"finished :)"