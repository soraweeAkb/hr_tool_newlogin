@echo off
if not exist main_online.exe exit 
echo 'Upgrading now...'
timeout /t 10 /nobreak
del main_app.exe
copy main_online.exe main_app.exe
del main_online.exe
echo 'Upgrade completed! Restarting...'
timeout /t 3 /nobreak
start main_app.exe
exit