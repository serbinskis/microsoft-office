@echo off

powershell -command "Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/serbinskis/microsoft-office/refs/heads/master/Install.bat' -OutFile \"$env:TEMP\Install.bat\""
cmd.exe /c "%TEMP%\Install.bat" & del "%TEMP%\Install.bat"

powershell -command "Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/serbinskis/microsoft-office/refs/heads/master/Activate.bat' -OutFile \"$env:TEMP\Activate.bat\""
cmd.exe /c "%TEMP%\Activate.bat" & del "%TEMP%\Activate.bat"
