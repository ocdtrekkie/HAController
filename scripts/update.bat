taskkill /im HAController.exe
C:\HAC\scripts\wget.exe -O C:\HAC\scripts\update.zip https://hac.jacobweisz.com/dl/update.zip
C:\HAC\scripts\7z.exe x -aoa -oC:\HAC\scripts\temp C:\HAC\scripts\update.zip
del /q C:\HAC\scripts\update.zip
robocopy C:\HAC\scripts\temp C:\HAC /E /R:6 /W:5
rmdir /s /q C:\HAC\scripts\temp
powershell -c echo `a
cd C:\HAC
C:\HAC\HAController.exe
exit
