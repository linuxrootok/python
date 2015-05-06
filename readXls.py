#!coding: utf-8
#!/bin/ent python
import types
import xlrd 
import xlwt

class read():
    '''
    read excel file
    '''
    def __init__(self,filename,sheetIndex=0,offset=1):
        self.filename = filename       
        self.offset = offset       
        book = xlrd.open_workbook(filename);
        sheet=book.sheets()[sheetIndex]
        self.sheetName = sheet
        self.nrows = sheet.nrows
        self.ncols = sheet.ncols
        self.lines = []
    def start(self):
        #读取数据
        for i in xrange(self.offset,self.nrows):
            lineData = [];
            for j in range(self.ncols):
                lineData.append(self.sheetName.cell(i,j).value)

            self.lines.append(lineData)
        return self.lines
    def createSql(self,table,col=[],filename=''):
        fp = open(filename,'w')
        sql = 'INSERT INTO'
        sql += ' `'+table+'` '
        sql += "('"+"','".join(col)+"')"
        sql += ' VALUES '
        for i in self.lines:
            x = []
            for j in i:
                if type(j) is types.StringType:
                    x.append(str(j.encode('utf-8')))
                elif type(j) is types.FloatType:
                    x.append(str(j))
                else:
                    x.append(j.encode('utf-8')) 
            sqlAll = sql+"('"+"','".join(x)+"');"
            fp.write(sqlAll+"\n")
            #print sqlAll
        fp.close()
        #print sql

def main():
    xls = read('test.xls',0,0)
    xls.start()
    xls.createSql('ttt',['a','b','c'],'aaa.sql')

if __name__ == '__main__':
    main()
