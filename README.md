## Install Microsoft Office 2024

Run the command line as administrator and execute the code below to install Microsoft Office.

```sh
powershell -command "Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/serbinskis/microsoft-office/refs/heads/master/Install.bat' -OutFile \"$env:TEMP\Install.bat\""
cmd.exe /c "%TEMP%\Install.bat" & del "%TEMP%\Install.bat"
```

## Activate Microsoft Office 2024

Run the command line as administrator and execute the code below to activate Microsoft Office.

```sh
powershell -command "Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/serbinskis/microsoft-office/refs/heads/master/Activate.bat' -OutFile \"$env:TEMP\Activate.bat\""
cmd.exe /c "%TEMP%\Activate.bat" & del "%TEMP%\Activate.bat"
```

## Install & Activate Microsoft Office 2024

Run the command line as administrator and execute the code below to install & activate Microsoft Office.

```sh
powershell -command "Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/serbinskis/microsoft-office/refs/heads/master/Setup.bat' -OutFile \"$env:TEMP\Setup.bat\""
cmd.exe /c "%TEMP%\Setup.bat" & del "%TEMP%\Setup.bat"
```
