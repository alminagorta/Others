# -*- coding: utf-8 -*-
"""
Created on Thu Jul 11 12:18:28 2019
This code was prepared to spare data from HTML to excel file using Python 3.7

inputs: html files
outputs :excel files

How to use?
1.Just put the html files in the same folder of this python code 
2.Create a folder "Output" and set the output path "out_path" in line 60
3. Excel files will be saved in this folder "Output"

Notes:
    HTM files's names must be <= 31 chars

@author: Omar Alminagorta
"""
#Install libraries
from bs4 import BeautifulSoup
import csv
import glob, os
import pandas as pd

#listing all files htm files
file1x= glob.glob("*.htm")#to list the files present
print(file1x)


# Loop to process all htm file=> From Htm to CSV file with space in rows
for file_to_work in glob.glob("*.htm"):
#    file_to_work= "ChloCreditRiveratCataractFalls.htm"
    html = open(file_to_work).read()
    soup = BeautifulSoup(html)
    table = soup.find("table")
    
    output_rows = []
    for table_row in table.findAll('tr'):
        columns = table_row.findAll('td')
        output_row = []
        for column in columns:
            output_row.append(column.text)
        output_rows.append(output_row)
    #writing csv file but with space
    nameCSV=file_to_work[11:-4]+'.csv'
    with open(nameCSV, 'w') as csvfile:#In Python 2.7. use wr instead of w
        writer = csv.writer(csvfile)
        writer.writerows(output_rows)

    
    #=========== read the csv file and passing to Excel to eliminate empty row==========  
    file1=pd.read_csv(nameCSV)    
    
    #writing Excel
    def saveExcel(fileX,nameX,sheetX):
        '''To save a fileX to XLSX in specific folder with nameX of the output and sheetX name
        e.g.=> saveExcel(benthicX,'try3.xlsx','aa')  '''
        out_name = nameX #set the name of the file
        #Set your output path after create your folder "Output"
        out_path="C:/Omar/BackUp_Nov18_2015/Toronto/Students_Feedback/Lauren_data/Output/"
        out_path = (out_path + out_name)
        writer = pd.ExcelWriter(out_path , engine='xlsxwriter')
        fileX.to_excel(writer, sheet_name=sheetX, header=False, index=False)
        writer.save()
        writer.close()
    
    
    saveExcel(file1,file_to_work[11:-4]+'.xlsx',file_to_work[11:-4])  