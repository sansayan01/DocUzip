@echo off
powershell -command "& {Start-Process cmd -ArgumentList '/c python your_script.py' -WindowStyle Minimized}"
