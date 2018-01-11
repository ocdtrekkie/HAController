taskkill /im HAController.exe
robocopy C:\HAC\scripts\temp C:\HAC /E /R:6 /W:5
rmdir /s /q C:\HAC\scripts\temp
powershell -c echo `a
cd C:\HAC
C:\HAC\HAController.exe
exit
