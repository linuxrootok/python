#!coding: utf-8
#!/bin/ent python
import xlrd 
import xlwt

class create():
    '''
    create excel file
    '''
    def __init__(self,filename,sheetName=[]):
        self.filename = filename       
        self.sheetName = sheetName
    def start(self):
        newbook = xlwt.Workbook()
        self.newbook = newbook
        if not self.sheetName:
            newsheet = newbook.add_sheet('sheet1')
            self.newsheet = newsheet
            self.isDefault = 1 
        else:
            self.isDefault = 0 

    def insert(self,lineArr=[]):
        for x,i in enumerate(lineArr):
            for y,z in enumerate(i):
                if self.isDefault:
                    self.newsheet.write(x,y,z)  
                else:
                    pass

        self.newbook.save(self.filename)    
def main():
    sheet = []
    xls = create('test.xls',sheet)
    xls.start()
    data = [[u'呵呵',2,3],[4,5,6,u'顶戴是速递ikk']]
    xls.insert(data)

if __name__ == '__main__':
    main()
