@echo off 

FOR %%i IN (*.docx) DO set foo=%%i


Set 7zip="C:\Program Files\7-Zip\7z.exe"

FOR /f "tokens=1 delims=." %%a IN ("%foo%") do set folder=%%a



IF EXIST %folder%\NUL goto compress ELSE goto extract




:extract
"C:\Program Files\7-Zip\7z.exe"  x %foo% -o* -aoa
goto end 


:compress
echo "Compress" 

"C:\Program Files\7-Zip\7z.exe"  a  %foo% .\%folder%\*
RMDIR /S /Q %folder%

goto end 


:end 
echo "end"