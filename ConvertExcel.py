from distutils import filelist
import string
import pandas as pd
import sys
import os
import configparser as cf


#Extract the excel files from the path
file_list=os.listdir(r"C:\Users\avinashv\myproject\pythonpoc")
#print (file_list)
fileList = []

print("extract the file names as tags for param config file ")
for l in file_list:
    #print(l)
    str1 = l
    if (str1[-5:] == '.xlsx'):
        fileList.append(str1[:-5])
#print(fileList) 

#print("Excel file names as config names") 
#fileList = ['test1', 'test2']


for f in fileList:

  #print (f)
  #Read param.txt file
  parser = cf.ConfigParser()
  parser.read("./param.txt")

  fileRef = f 

  print(parser.get(fileRef,'filenm'))
  n = int(parser.get(fileRef, "cnt"))
  #print(n)

  for x in range(1,n+1):
    #print(x)
    cs = 'csv' + str(x)
    csvfile = parser.get(fileRef, cs)
    string = csvfile
    SheetNm  = string[0:-4]
    path = parser.get(fileRef, "path")
    excelFile = path + '\\' + fileRef +'.xlsx'
    
    #print(SheetNm)
    #print (excelFile) 
    #print(csvfile)

    #convert Excel sheet to csv file
    read_file = pd.read_excel (excelFile, sheet_name = SheetNm)
    read_file.to_csv (path + r'\\' + csvfile , index = None, header=True)


#convert xxxx sheet to csv file
#read_file = pd.read_excel (path + r'\xxxx.xlsx', sheet_name='xxxx')
#read_file.to_csv (path + r'\xxxx.csv', index = None, header=True)

print ('Successfully Csv files generated for all the Excels provided')

