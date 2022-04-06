Dim appExcel
dim wbk
dim sht

Set appExcel = WScript.CreateObject("Excel.Application")
appExcel.Visible = True
Set wbk = appExcel.Workbooks.Add()

set sht = wbk.Sheets(1)
sht.Range("A1") = "create_blank_workbook.vbs"
