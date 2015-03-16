echo %time%__call__transform >> log.txt
echo input_file_%1 >> log.txt
echo output_file_%2 >> log.txt

rem gswin32 -dSAFER -dBATCH -dNOPAUSE -sDEVICE=jpeg -r100  -sOutputFile=figure-%%03d.jpg 123.prn

cd /d %~dp0
cd ghostScript
gs -dSAFER -dBATCH -dNOPAUSE -sDEVICE=pdfwrite  -dEPSCrop -dCompatibilityLevel=1.4 -dGraphicsAlphaBits=1 -dTextAlphaBits=1  -r100 -sOutputFile=%2 %1 >>log.txt


forfiles /p %3 /d -1 /c "cmd /c del /f @path"




