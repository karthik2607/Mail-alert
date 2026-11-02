@echo off
rem Starts MailAlert without a console window. Stop it from the tray icon > Quit.
cd /d "%~dp0"
start "" pythonw mailalert.py
