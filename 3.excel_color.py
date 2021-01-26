import win32com.client

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Open('c:\\Users\\d\\Desktop\\test.xlsx')
ws = excel.ActiveSheet

ws.Cells(1,2).Value = "is"
ws.Range("C1").value = "good"
ws.Range("C1").Interior.ColorIndex = 10

