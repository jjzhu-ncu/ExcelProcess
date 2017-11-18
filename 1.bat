@echo off
echo hello
set curr_dir=%cd%
set ENV_PATH=%PATH%;%cd%\Python34
echo %ENV_PATH%
echo %curr_dir%
%cd%\Python34\python.exe test.py
pause