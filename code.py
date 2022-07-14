
# coding: utf-8

# In[1]:


names = list()                                                 #list of arrays used
ages = list()
genders = list()
phones = list()


# In[2]:


import pandas as pd
import xlsxwriter                               #importing the xlsxwriter package for further processing                  
import win32com.client as win32                              
import openpyxl
from xlrd import open_workbook

def write():
    workbook = xlsxwriter.Workbook('LIST1.xlsx')     #opening a excel workbook() in the name given inside the called fuction
    worksheet = workbook.add_worksheet()            #adding new worksheet to the workbook using add_worksheet() function

    cell1 = workbook.add_format()                   #add_format() is used to add a new cell format for styling purpose
    cell1.set_bold()                                #set_bold() : Turn on bold for the format font
    cell1.set_font_size(20)                         #set_font_size(size) : Set the size of the font used in the cell.

    cell2 = workbook.add_format()                   #adding another cell format with different name
    cell2.set_font_name('MONOTON')                  #set_font_name('style') : Specify the font used used in the cell.

    worksheet.write('A1', 'NAME',cell1)             #writing into the sheet using cell notation
    row = 1                                         #initializing the value of row to 1
    column = 0                                      #initializing the value of column to 0
    for item in names :                             #running for loop till all the name are written into the sheet
        worksheet.write(row, column, item)          #write the values from the array declared above one by one into each cell
        row += 1                                    #incrementing the value of row,  row=row+1

    worksheet.write('B1', 'AGE',cell2)             
    row = 1
    column=1
    for item in ages : 
        worksheet.write(row, column, item) 
        row += 1
    
    worksheet.write('C1', 'GENDER')
    row = 1
    column=2
    for item in genders : 
        worksheet.write(row, column, item) 
        row += 1
    
    worksheet.write('D1', 'PHONE NUMBER')
    row = 1
    column=3
    for item in phones : 
        worksheet.write(row, column, item) 
        row += 1
    workbook.close() #closing the workbook after completion of procedure                       

    
def autofit():                 #this whole set of code is used to autofit the size of each cell asper the size of the input data
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open('C:\Users\Abhishek\LIST1.xlsx')
        ws = wb.Worksheets("Sheet1")
        ws.Columns.AutoFit()
        wb.Save()
        excel.Application.Quit()

def intro():
    print '----------------------------------------------------------------------------------------'
    print 'Welcome'
    print 'The things u can do here are:'
    print '1.Adding value to the excel file'
    print '2.Deleting a perticular detail'
    print '3.Print a persons detail'
    print '4.Print a column'
    print '5.Exit function'
    arg= raw_input("Enter the operation you want to perform:")

    if arg=='1':
        add_val()
    elif arg=='2':
        del_val()
    elif arg=='3':
        row_dis()
    elif arg=='4':
        col_dis()
    else:
        print "Thank you for using this"


def add_val():
    num = raw_input("Enter how many entries u wanna make:")    #using raw_input() for small inputs
    print 'Enter the list: '
    for i in range(int(num)):
        print ("\nDETAIL %d" %(i+1))
        name = raw_input("name :")                             #defining another variable for inputing the values
        names.extend([name])                                #adding the values to the array
        age = raw_input("age :")
        ages.extend([int(age)])
        gender = raw_input("gender(M\F\T) :")
        genders.extend([gender])
        phone = raw_input("phone number :")
        phones.extend([int(phone)])
    write()
    autofit()
    intro()

        
def del_val():
    name=raw_input("Enter the name you want to delete:")
    i=0
    while i < len(names):
        if names[i] == name:
            del names[i]
            del ages[i]
            del genders[i]
            del phones[i]
        else:
            i += 1
    write()
    autofit()
    intro()
        
def row_dis():
    book = open_workbook("C:\Users\Abhishek\LIST1.xlsx")
    i=raw_input("Enter the name whose details you want to see:")
    for sheet in book.sheets():
        for rowidx in range(sheet.nrows):
            row = sheet.row(rowidx)
            for colidx, cell in enumerate(row):
                if cell.value == i :
                    print '-------------------------------------'
                    print sheet.row_values(rowidx)
    intro()
    
def col_dis():
    print 'Column u can see are:'
    print '1.Name'
    print '2.Age'
    print '3.Gender'
    print '4.Phone Numbers'
    i=raw_input("Enter the column name u want to see:")
    print '--------------'
    if i=='1':
        for x in range(len(names)): 
            print names[x] 
    elif i=='2':
        for x in range(len(ages)): 
            print ages[x]
    elif i=='3':
        for x in range(len(genders)): 
            print genders[x]
    elif i=='4':
        for x in range(len(phones)): 
            print phones[x]
    else:
        print'No such column exist'
    intro()


# In[3]:


intro()

