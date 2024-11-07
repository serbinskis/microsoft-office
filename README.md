## Install Microsoft Office 2016

Run the command line as administrator and execute the code below to install Microsoft Office 2016.

```sh
powershell -Command "Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/serbinskis/microsoft-office-2016/refs/heads/master/Install.bat' -OutFile \"$env:TEMP\Install.bat\""
cmd.exe /c "%TEMP%\Install.bat"
powershell -Command "Remove-Item \"$env:TEMP\Install.bat\""
```

## Activate Microsoft Office 2016

Run the command line as administrator and execute the code below to activate Microsoft Office 2016.

```sh
powershell -Command "Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/serbinskis/microsoft-office-2016/refs/heads/master/Activate.bat' -OutFile \"$env:TEMP\Activate.bat\""
cmd.exe /c "%TEMP%\Activate.bat"
powershell -Command "Remove-Item \"$env:TEMP\Activate.bat\""
```

## Install & Activate Microsoft Office 2016

Run the command line as administrator and execute the code below to install & activate Microsoft Office 2016.

```sh
powershell -Command "Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/serbinskis/microsoft-office-2016/refs/heads/master/Setup.bat' -OutFile \"$env:TEMP\Setup.bat\""
cmd.exe /c "%TEMP%\Setup.bat"
powershell -Command "Remove-Item \"$env:TEMP\Setup.bat\""
```
