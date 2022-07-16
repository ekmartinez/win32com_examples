# -*- coding: utf-8 -*-
"""
Created on Thu Jul 26 08:20:21 2018

@author: CDRMartiE
"""

from win32com.client import Dispatch

#Application
xl = Dispatch("Excel.Application")
wb = xl.Workbooks.Open("N:\\Book.xls")
xl.Visible = 1
ws = wb.Sheets("Sheet1")

#CopyPaster
ws.Range("A5:A5000").Copy()
ws.Paste(ws.Range('A6'))

#row deleter
ws.Range('A1:A2').EntireRow.Delete()
ws.Range('A2').EntireRow.Delete()

#column deleter
ws.Range('B1').EntireColumn.Delete()
ws.Range('F:H').EntireColumn.Delete()

#clear contents
ws.Range('A2').ClearContents()

columns = ['Invoice Number', 'Invoice Date', 'Payment Date', 'Payment Number', 'Amount']
ws.Range('A3:E3').value = columns
        
#Loop
rng_del = ws.Range('E4:E5000')
for i in rng_del:
    if i.Value == None:
        i.EntireRow.Delete()
        
for x in rng_del:
    if x.Value == None:
        x.EntireRow.Delete()    
        
#autofit
ws.Columns.AutoFit()

#alignment
xlLeft, xlRight, xlCenter = -4131, -4152, -4108 #Use the one you need.
ws.Range("A:E").HorizontalAlignment = xlCenter    

#XlDirectionDown = 4
#last = wb.Range("A5:A5").End(XlDirectionDown)
        