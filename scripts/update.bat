taskkill /im HAController.exe
C:\HAC\scripts\wget.exe -O C:\HAC\scripts\update.zip http://files.ocdtrekkie.com/dl/hac/update.zip
C:\HAC\scripts\7z.exe x -aoa -oC:\HAC\scripts\temp C:\HAC\scripts\update.zip
del /q C:\HAC\scripts\update.zip
robocopy C:\HAC\scripts\temp C:\HAC /E
rmdir /s /q C:\HAC\scripts\temp
cd C:\HAC
C:\HAC\HAController.exe
exit