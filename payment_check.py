import openpyxl
from openpyxl import load_workbook
keywordlist = ['山语墅','西区供水','undefined']
class projectinfo:
    name='undefined'
    def _init_(self,serialnum,price,location):
        self.serialnum = serialnum
        self.price = price
        self.location = location
    def keywordcheck(selfname):
        for i in keywordlist:
            if selfname.find(i):
                self.name=i
    collectioninfo = 0
    specimentype = ''
    
    pass

if __name__ == '__main__':
    #read xlsx file as exported and collect basic datas for future use
    wb =load_workbook(filename = 'D:/pt/registrationexport.xlsx')
    wb.active
    ws = wb.get_sheet_by_name('Sheet1')
    for i in range(3,ws.max_row):
        proj = projectinfo(ws.cell(i,2).value,ws.cell(i,15).value,ws.cell(i,12).value)
        proj.keywordcheck(ws.cell(i,13).value)
        
