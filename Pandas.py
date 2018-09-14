import numpy as np
import pandas as pd
import openpyxl
import xlsxwriter
// Start
def main():
    df,file1 = getSourceFile()
    fileSh = getSheet(df,file1)
    print(fileSh.columns.values)
    sColumn,column1 = getScolumn(fileSh)
    df2,file2 = getCompareFile()
    fileSh2 = getSheet(df2,file2)
    cColumn,column2 = getCcolumn(fileSh2)
    dfnew = pd.merge(sColumn, cColumn, how='outer', right_on=column2, left_on=column1,indicator="Found In",validate="m:m",suffixes=("_"+file1,"_"+file2),sort=True)
    dfnew = dfnew.replace('left_only','Not in '+file2, regex=True)
    dfnew = dfnew.replace('right_only', 'Not in ' + file1, regex=True)
    nameFile = input("Enter new file name: ")
    writer = pd.ExcelWriter(nameFile+".xlsx", engine='xlsxwriter')
    dfnew.to_excel(writer, 'Sheet1')
    writer.save()
    print("Done")
def getSourceFile():

    while True:
        try:
            sname = input("Enter Source File Name: ")
            df = pd.ExcelFile(sname)

        except FileNotFoundError:
            print("File not found")
            continue
        else:
            return df,sname
            break
def getCompareFile():
    while True:
        try:
            name = input("Enter Compare File Name: ")
            sname = name
            df2 = pd.ExcelFile(sname)
        except FileNotFoundError:
            print("File not found")
            continue
        else:

            return df2,sname
            break

def getScolumn(df):
    test = df.columns.values
    for i in range(len(test)):
        print(str(i + 1) + ":", test[i])
    while True:
        try:
            option = int(input("Please enter number of column you wish to work on: "))
            column1 = test[option - 1]
            sColumn = df.loc[:, [column1]]
        except:
            print("Out of bound, please enter the number of the column")
            continue
        else:
            return sColumn, column1
def getCcolumn(df2):
    test = df2.columns.values
    for i in range(len(test)):
        print(str(i + 1) + ":", test[i])
    while True:
        try:
            option = int(input("Please enter number of column you wish to work on: "))
            column1 = test[option-1]
            sColumn = df2.loc[:,[column1]]
        except:
            print("Out of bound, please enter the number of the column")
            continue
        else:
            return sColumn, column1


'''
    while True:
        try:
            SourceColumn = input("Enter Compare Column: ")
            sColumn = df2.loc[:, [SourceColumn]]
        except KeyError:
            print("Column not found")
            continue
        else:
            return sColumn,SourceColumn
            break
'''

def getSheet(df,file1):
    while True:
        try:
            options = df.sheet_names
            for i in range(len(options)):
                print(str(i + 1) + ":", options[i])

            option = int(input("Please enter the number of sheet to work on: "))
            shName = options[option-1]

            fileSheet = pd.read_excel(file1, sheet_name= shName)
        except:
            print("No such sheet exists")
            continue
        else:
            return fileSheet
            break



main()
