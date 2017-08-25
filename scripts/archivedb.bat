rem https://stackoverflow.com/questions/18895422/how-to-use-a-batch-file-to-rename-a-file-to-include-the-date/18907550#18907550

taskkill /im HAController.exe
timeout /t 5 /nobreak

for /f "delims=" %%a in ('wmic OS Get localdatetime  ^| find "."') do set "dt=%%a"
set "YY=%dt:~2,2%"
set "YYYY=%dt:~0,4%"
set "MM=%dt:~4,2%"
set "DD=%dt:~6,2%"
set "HH=%dt:~8,2%"
set "Min=%dt:~10,2%"
set "Sec=%dt:~12,2%"

set datestamp=%YYYY%%MM%%DD%
set timestamp=%HH%%Min%%Sec%
set fullstamp=%YYYY%-%MM%-%DD%_%HH%-%Min%-%Sec%

md "C:\HAC\archive"
copy "C:\HAC\HACdb.sqlite" "C:\HAC\archive\HACdb_%fullstamp%.sqlite"

powershell -c echo `a
cd C:\HAC
C:\HAC\HAController.exe
exit