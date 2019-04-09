#usr/bin/python
# -*-coding:  utf-8 -*-
import xlwt,xlrd,sys
import re


def open_excel(file= 'SPRDLogDatabase_V1.0.xlsx'):
    try:
        data = xlrd.open_workbook(file)
        return data
    except:
        print ("error")


def excel_table(file="SPRDLogDatabase_V1.0.xlsx",colnameindex=4, by_name=u'Sheet1'):
    data=open_excel(file)  #打开excel
    table=data.sheet_by_name(by_name) #根据sheet名字来获取excel中的sheet
    nrows=table.nrows  #获取行数
    ncols=table.ncols  #获取列数
    colnames=table.col_values(colnameindex)#第一行数据
    print ("colname===",colnames)
    list=[]#放入读取结果
    for rownum in range(0,nrows): #遍历每一行
        print ("rownum==",rownum)  #获取某一行
        row=table.row_values(rownum)
        if row:
            app=[]
            for i in range(len(colnames)):#一列一列读取
                app.append(row[i])
            list.append(app)
    return list
            

def seek(seek_name):
    suggestions=[]
    pattern = '.*'.join(seek_name)
    regex=re.compile(pattern)
    
    
    

def operta_io(tables):
    file=open("1.txt","a")
    for row in tables:
        print ("row===",row)
        file.write(str(row)+'\n') #将excel里边的内容写入到文本文件并进行换行
    file.close()


def main():
    tables=excel_table()
    operta_io(tables)
    





if __name__=="__main__":
    main()
