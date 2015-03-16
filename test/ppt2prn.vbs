' Automate PowerPoint to print a document to activePDF Server
Set objPPT = CreateObject("PowerPoint.Application")
Set objDoc = objPPT.Presentations.Open("c:\4.pptx", -1, 0, 0)
Set objOptions = objDoc.PrintOptions
objOptions.ActivePrinter =  "InfoView Document Printer"
objOptions.PrintInBackground = 0
objDoc.PrintOut 1, 9999, "c:\ppt.prn", 1, 0
objDoc.Saved = 1
objDoc.Close
objPPT.Quit
Set objOptions = Nothing
Set objDoc = Nothing
Set objPPT = Nothing