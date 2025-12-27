@echo off
set LOGDIR=%LOCALAPPDATA%\save-contacts-outlook
set LOGFILE=%LOGDIR%\run.log
python "%USERPROFILE%\save-contacts-outlook\process_email.py" "%~1" >> "%LOGFILE%" 2>&1
