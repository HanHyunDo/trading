import win32com.client

class CpStockCode:
    def __init__(self):
        self.stocks = {'유한양행':'A000100'}
    def GetCount(self):
        return len(self.stocks)
    def NameToCode(self, name):
        return self.stocks[name]

instCpStockCode = CpStockCode()

print(instCpStockCode.GetCount())
print(instCpStockCode.NameToCode('유한양행'))

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Add()
ws = wb.Worksheets("Sheet1")
ws.Cells(1, 1).Value = "hello world"
wb.SaveAs('c:\\Users\\d\\Desktop\\test.xlsx')
excel.Quit()