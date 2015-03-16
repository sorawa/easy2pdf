Set fso = Wscript.CreateObject("Scripting.FileSystemObject")
set c=fso.getfile("C:\WINDOWS\MODIVCDemo.exe") '被拷贝的文件的位置
c.copy("拷贝来的注册表编辑器.exe") '拷贝到哪(可以是绝对路径,这里是拷贝到运行的*.vbs文件所在文件夹!