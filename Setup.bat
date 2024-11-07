@echo off

powershell -Command "Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/serbinskis/microsoft-office-2016/refs/heads/master/Install.bat' -OutFile \"$env:TEMP\Install.bat\""
cmd.exe /c "%TEMP%\Install.bat"
powershell -Command "Remove-Item \"$env:TEMP\Install.bat\""

powershell -Command "Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/serbinskis/microsoft-office-2016/refs/heads/master/Activate.bat' -OutFile \"$env:TEMP\Activate.bat\""
cmd.exe /c "%TEMP%\Activate.bat"
powershell -Command "Remove-Item \"$env:TEMP\Activate.bat\""
