rem gswin32 -dSAFER -dBATCH -dNOPAUSE -sDEVICE=jpeg -r100  -sOutputFile=figure-%%03d.jpg 123.prn

cd ghostScript
gs -dSAFER -dBATCH -dNOPAUSE -sDEVICE=pdfwrite  -dEPSCrop -dCompatibilityLevel=1.4 -dGraphicsAlphaBits=1 -dTextAlphaBits=1  -r100 -sOutputFile=../output.pdf  ../out.prn


echo %time% >> log.txt


