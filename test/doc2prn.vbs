

Set objWD = CreateObject("Word.Application")
Set objDoc = objWD.Documents.Open("D:\oacn\htdocs\oa.cn\protected\script\pdftsfm\doc_file\1.doc")

objWD.ActivePrinter = "InfoView Document Printer"
'objWD.ActiveWindow.ActivePane.VerticalPercentScrolled = 21
'objWD.ActiveWindow.ActivePane.VerticalPercentScrolled = 21
'objWD.ActiveWindow.ActivePane.VerticalPercentScrolled = 50

'Range:=wdPrintAllDocument, Item:= _
'        wdPrintDocumentWithMarkup, Copies:=1, Pages:="", PageType:= _
'        wdPrintAllPages, ManualDuplexPrint:=False, Collate:=True, Background:= _
'        True, PrintToFile:=True, PrintZoomColumn:=0, PrintZoomRow:=0, _
'        PrintZoomPaperWidth:=0, PrintZoomPaperHeight:=0, OutputFileName:="", _
'        Append:=False


file_path = CreateObject("Scripting.FileSystemObject").GetFolder(".").Path
file_path = file_path & "\" & "out.prn"

objWD.PrintOut 1,0,0,file_path,,,7,1


objDoc.Close
set objWD=nothing 




