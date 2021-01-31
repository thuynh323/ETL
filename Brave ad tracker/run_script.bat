@echo off
:Ask
echo Sorry...Have I run today? Please check the database if you don't remember. (y/n)
set INPUT=
set /P INPUT=Type input: %=%
If /I "%INPUT%"=="y" goto yes 
If /I "%INPUT%"=="n" goto no
echo Incorrect input & goto Ask

:yes
echo You should get some sleep...
pause
exit

:no
echo Start the program...
cd C:\Users\Thanh Huynh\Documents\Projects\brave-ads\venv
"C:\Python38\python.exe" "C:\Users\Thanh Huynh\Documents\Projects\brave-ads\venv\main.py"
pause
