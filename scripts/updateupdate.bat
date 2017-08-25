C:\HAC\scripts\wget.exe -O C:\HAC\scripts\updateupdate.zip https://hac.jacobweisz.com/dl/updateupdate.zip
C:\HAC\scripts\7z.exe x -aoa -oC:\HAC\scripts\temp C:\HAC\scripts\updateupdate.zip
del /q C:\HAC\scripts\updateupdate.zip
robocopy C:\HAC\scripts\temp C:\HAC\scripts /E
rmdir /s /q C:\HAC\scripts\temp
powershell -c echo `a
exit
