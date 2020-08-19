import xml.etree.ElementTree as ET
from glob import glob
import os
import sys
import xlwt 
from xlwt import Workbook


def selectAllFiles():

#    os.chdir(r"C:\Users\enkisum\OneDrive - Ericsson AB\Optic\XMLParser\Test")
    os.chdir(sys.argv[1])
    global files
    files = glob("*.xml")
    return files

def writetoExcel():
   
    wb = Workbook() 
    sheet1 = wb.add_sheet('Sheet 1')
    x=0
    global root
    
    for file in files:
        x=x+1
        with open(file,"rb") as data:
            cta = ET.parse(data)
#            print(cta)
            root = cta.getroot()
#            print(root)
            for i in root.findall('filename'):
                sheet1.write(x, 0, i.text)
#                print("filename:"+i.text)
            for j in root.findall('path'):
#                print("path:")
#                print(j.text)
                sheet1.write(x, 1, os.path.join("path:",str(j.text)))
#                print(os.path.join("path:",str(j.text)))
            for k in root.findall('object'):
                sheet1.write(x, 2, k.find('class').text)
                sheet1.write(x, 3, k.find('name').text)
#                print("class:"+k.find('class').text)
#                print("label:"+k.find('name').text)
                
        
    wb.save('validatexml.xls')

def parseData():
   
    global root
    for file in files:
        with open(file,"rb") as data:
            cta = ET.parse(data)
#            print(cta)
            root = cta.getroot()
#            print(root)
            for i in root.findall('filename'):
                print("filename:"+i.text)
            for j in root.findall('path'):
#                print("path:")
#                print(j.text)
                print(os.path.join("path:",str(j.text)))
            for k in root.findall('object'):
                print("class:"+k.find('class').text)
                print("label:"+k.find('name').text)

                   	    
selectAllFiles()
parseData()
writetoExcel()