Function  ExcelRead(xlpath,xlsheet,row,col)

Dim mysheet
Dim myxlapp

Set myxlapp=createobject("Excel.Application")
myxlapp.Workbooks.Open xlpath
Set mysheet=myxlapp.ActiveWorkbook.Worksheets(xlsheet)

ExcelRead=mysheet.Cells(row,col)

myxlapp.ActiveWorkbook.Close
myxlapp.Application.Quit

Set myxlapp=Nothing
Set mysheet=Nothing
End Function

Function  ExcelWrite(xlpath,xlsheet,row,col,xldata)

Dim mysheet
Dim myxlapp

Set myxlapp=createobject("Excel.Application")
myxlapp.Workbooks.Open xlpath
Set mysheet=myxlapp.ActiveWorkbook.Worksheets(xlsheet)

mysheet.Cells(row,col)=xldata
myxlapp.ActiveWorkbook.Save

myxlapp.ActiveWorkbook.Close
myxlapp.Application.Quit

Set myxlapp=Nothing
Set mysheet=Nothing

End Function
