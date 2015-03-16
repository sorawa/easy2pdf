On Error Resume Next
Set objApp = CreateObject("Excel.Application")
Set objDoc = objApp.Workbooks.Open ("c:\3.xls")

'objApp.ActivePrinter = "InfoView Document Printer"
'objApp.ActiveSheet.PageSetup.PaperSize = 8
'objApp.ActiveSheet.PageSetup.PrintArea = "A1:AD49"

cols = objApp.ActiveSheet.UsedRange.Columns.Count

if cols < 10 then
	objApp.ActiveSheet.PageSetup.PaperSize = 9
	objApp.ActiveSheet.PageSetup.Zoom = 60 
	
end if

if cols >= 10 and cols =<20 then
	objApp.ActiveSheet.PageSetup.PaperSize = 172
	objApp.ActiveSheet.PageSetup.Zoom = 55 
	
end if

if cols >= 20 and cols =<30 then
	objApp.ActiveSheet.PageSetup.PaperSize = 180
	objApp.ActiveSheet.PageSetup.Zoom = 45 
	
end if

if cols > 30 and cols =<50 then
	objApp.ActiveSheet.PageSetup.PaperSize = 180
	objApp.ActiveSheet.PageSetup.Zoom = 40 
	
end if

file_path = CreateObject("Scripting.FileSystemObject").GetFolder(".").Path
file_path = file_path & "\" & "out.prn"

objApp.Sheets.PrintOut ,,1,0,"InfoView Document Printer",1,True,file_path,1


objDoc.Close True
set objApp=nothing 

err.Clear