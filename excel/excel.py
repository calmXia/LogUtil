#! python3
#! usr/bin/python
# -*-coding:  utf-8 -*-
import xlwt,xlrd,sys
import re

#LogDataBase = r'C:\xxx\SPRDLogDatabase_V1.1.xlsx'
# or
#LogDataBase = 'C:/xxx/SPRDLogDatabase_V1.1.xlsx'

def cprint():
    return print("%s.%s: start_rowx %d end_rowx %d -->" % (self.__class__.__name__, sys._getframe().f_code.co_name, start_rowx, end_rowx))


class Excel(object):

    '''
    methods:
        __init__
        
        open
        
        init_table
        
        getdata_by_col_index
        
        getdata_by_module
        
        
    '''

    # constructor
    def __init__(self, file):
        self.__file = file
        self.__title = (
                        "ARCH_Component", "Module",      "Component", 
                        "LogType",	      "LogText",     "LogFile", 
                        "Analization",    "Root-Cause",  "Resolution",
                        "Suggestion",     "SW-Platform", "CodeFile",
                        "CQ"
                        )
        self.__nrows = 0
        self.__ncols = 0
        self.__table = self.init_table(workbook = self.open())
        # set title by tuple

    
    def open(self):
        try:
            workbook = xlrd.open_workbook(self.__file)
            return workbook
        except FileNotFoundError as e:
            print("FileNotFoundError: ", e)

    def init_table(self, workbook, name=u'Sheet1'):
        print('init_table...')
        table = workbook.sheet_by_name(name) #根据sheet名字来获取excel中的sheet
        
        self.__nrows = table.nrows  #获取有效行数
        self.__ncols = table.ncols  #获取有效列数
        print('Table %s nrows %d ncols %d' % (name, self.__nrows, self.__ncols))
        print("Title : %s " % (self.__title, ))
        return table
        
        '''
        colnames=table.col_values(colnameindex)#第一行数据
        print ("colname===",colnames)
        #list=[]#放入读取结果
        for rownum in range(0,nrows): #遍历每一行
            print ("rownum==",rownum)  #获取某一行
            row=table.row_values(rownum)
            if row:
                col_collections=[]
                for i in range(len(colnames)):#一列一列读取
                    col_collections.append(row[i])
            else:
                return col_collections
        '''    
        
    def get_index_by_title(self, module=None):
        print("start to found module -- %s " %(module))
        for i in self.__title:
            if self.__title[i] is module:
                return i
            else:
                print("not found %s" % (module))
        
    # Gets the specified column
    # Func: table.col_values(colx, start_rowx=0, end_rowx=None)
    # 返回由该列中所有单元格的数据组成的列表
    # Note: start_rowx=1 : remove column title
    def getdata_by_col_index(self, col_index, start_rowx=1, end_rowx=None):
        print("%s.%s: start_rowx %d end_rowx %d -->" % 
            (self.__class__.__name__, sys._getframe().f_code.co_name, start_rowx, end_rowx))
        return self.__table.col_values(col_index, start_rowx, end_rowx)

    def getdata_by_module(self, module=None):
        index = get_index_by_title(module)
        #start_row = getdata_by_col_index()
        return 

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


# test
'''
def main():
    #tables=excel_table()
    LogDataBase=sys.argv[1]
    tables = Excel(LogDataBase)
    print(tables.getdata_by_col_index(4, end_rowx = 4))
    #workbook = tables.open()
    #tables.excel_table(workbook)
    #operta_io(tables)
    





if __name__=="__main__":
    main()
'''
