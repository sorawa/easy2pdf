'********************************************
'Main
'*******************************************
if WScript.Arguments.Count < 2 then 
	wscript.quit
end if

input_file = WScript.Arguments(0) 
ouput_file = WScript.Arguments(1) 
'copy to temp file
f0 = WScript.ScriptFullName
Set fso = CreateObject("Scripting.FileSystemObject")
abs_path = fso.GetAbsolutePathName(f0 & "\..")



temp_file = abs_path & "\temp_file\" & (Replace(Time(),":","")) & "_" & GetFileName(input_file)
prn_file =  abs_path & "\temp_file\" & (Replace(Time(),":","")) &  "_prn.prn"

CopyFileTo input_file,temp_file 

'choice the ext to transform
'make the ext into upper case
'dont open the source file , it make exception
'and we can open the temp file
file_ext = UCase(GetFileExt(input_file))
' WORD document 
if  file_ext = ".DOC" or file_ext = ".DOCX" then
	a = doc2prn(temp_file,prn_file)
end if
'Excel document
if  file_ext = ".XLS" or file_ext = ".XLSX" then
	a = xls2prn(temp_file,prn_file) 
end if
'PPT document 
if  file_ext = ".PPT" or file_ext = ".PPTX" then
	a = ppt2prn(temp_file,prn_file) 
end if

'after printer handle , we got the .prn file
'at last transform .prn to .pdf
Set shell=Wscript.createobject("wscript.shell")
ouput_file = replace(ouput_file,"\","/")
gscmd = abs_path & "\gs\transform_oncall.bat " & prn_file & " " & ouput_file & " " & abs_path & "\temp_file"
'msgbox gscmd
result = shell.run(gscmd,0)
'msgbox gscmd
'MsgBox GetFilePath(input_file)
'MsgBox GetFileExt(input_file)
'MsgBox GetFileName(input_file)

'(Int(Rnd()*100000) & Int(Rnd()*100000))
'sleep(1000)
'********************************************
'func
'c:\abc\566\123.txt  c:\abc\566\456.txt
'*******************************************
Function CopyFileTo(f_from,f_to)
	Set fso = Wscript.CreateObject("Scripting.FileSystemObject")
	set c=fso.getfile(f_from) 	'to copy
	c.copy(f_to) 				'to where
End Function

'********************************************
'func
'c:\abc\566\123.txt  return c:\abc\566\
'*******************************************
Function GetFilePath(fp)
	n = 0
	ret = 0
	do
		'msgbox(Right(fp,n))
		if Left(Right(fp,n),1) = "\" then
			ret = 1
		else 
			n=n+1 
		end if
	loop until n = len(fp) - 1  or ret = 1
	GetFilePath = Left(fp,len(fp) - n + 1)
	
End Function

'********************************************
'func
'c:\abc\566\123.txt  return .txt
'*******************************************
Function GetFileExt(fp)
	n = 0
	ret = 0
	do
		'msgbox(Right(fp,n))
		if Left(Right(fp,n),1) = "." then
			ret = 1
		else 
			n=n+1 
		end if
	loop until n = len(fp) - 1  or ret = 1
	GetFileExt = Right(fp,n)
End Function

'********************************************
'func
'c:\abc\566\123.txt  return 123.txt
'*******************************************
Function GetFileName(fp)
	n = 0
	ret = 0
	do
		'msgbox(Right(fp,n))
		if Left(Right(fp,n),1) = "\" then
			ret = 1
		else 
			n=n+1 
		end if
	loop until n = len(fp) - 1  or ret = 1
	GetFileName = Right(fp,n-1)
End Function


'********************************************
'func
'c:\abc\566\123.txt  return 123.txt
'*******************************************
Function xls2prn(xls_file,prn_file)
	On Error Resume Next
	Set objApp = CreateObject("Excel.Application")
	Set objDoc = objApp.Workbooks.Open (xls_file)
	
	cols = objApp.ActiveSheet.UsedRange.Columns.Count
	
	if cols < 10 then
		objApp.ActiveSheet.PageSetup.PaperSize = 9
		objApp.ActiveSheet.PageSetup.Zoom = 60 
	end if

	if cols >= 10 and cols =<20 then
		objApp.ActiveSheet.PageSetup.PaperSize = 9
		objApp.ActiveSheet.PageSetup.Zoom = 60 
	end if

	'msgbox cols
	
	if cols >= 20 and cols =<30 then
		objApp.ActiveSheet.PageSetup.PaperSize = 8
		objApp.ActiveSheet.PageSetup.Zoom = 50 
	end if

	if cols > 30 and cols =<50 then
		objApp.ActiveSheet.PageSetup.PaperSize = 180
		objApp.ActiveSheet.PageSetup.Zoom = 120 
	end if

	'file_path = CreateObject("Scripting.FileSystemObject").GetFolder(".").Path
	'file_path = file_path & "\" & "out.prn"

	objApp.Sheets.PrintOut ,,1,,"InfoView Document Printer",1,True,prn_file,0


	objDoc.Close True
	set objApp=nothing 

	err.Clear
End Function

'********************************************
'func
'c:\abc\566\123.txt  return 123.txt
'*******************************************
Function doc2prn(doc_file,prn_file)
	Set objWD = CreateObject("Word.Application")
	Set objDoc = objWD.Documents.Open(doc_file)
	objWD.ActivePrinter = "InfoView Document Printer"

	'file_path = CreateObject("Scripting.FileSystemObject").GetFolder(".").Path
	'file_path = file_path & "\" & "out.prn"
	'MsgBox prn_file
	objWD.PrintOut 1,0,0,prn_file,,,7,1
	'objWD.PrintOut 1,0,0,prn_file,,,7,1,0,1,True,""

	objDoc.Close
	set objWD=nothing 

End Function


'********************************************
'func
'c:\abc\566\123.txt  return 123.txt
'*******************************************
Function ppt2prn(ppt_file,prn_file)
	Set objPPT = CreateObject("PowerPoint.Application")
	Set objDoc = objPPT.Presentations.Open(ppt_file, -1, 0, 0)
	Set objOptions = objDoc.PrintOptions
	
	objOptions.ActivePrinter =  "InfoView Document Printer"
	objOptions.PrintInBackground = 0
	objDoc.PrintOut 1, 9999, prn_file, 1, 0
	objDoc.Saved = 1
	objDoc.Close
	objPPT.Quit
	
	Set objOptions = Nothing
	Set objDoc = Nothing
	Set objPPT = Nothing
End Function




