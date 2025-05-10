@echo off
title Downloading - Microsoft Office 2016
pushd %~dp0

md "%TEMP%\Office2016"
cd "%TEMP%\Office2016"

powershell -Command "(New-Object Net.WebClient).DownloadFile('https://github.com/serbinskis/microsoft-office-2016/raw/refs/heads/master/Office2016/Office2016.zip.001', 'Office2016.zip.001')"
powershell -Command "(New-Object Net.WebClient).DownloadFile('https://github.com/serbinskis/microsoft-office-2016/raw/refs/heads/master/Office2016/Office2016.zip.002', 'Office2016.zip.002')"
powershell -Command "(New-Object Net.WebClient).DownloadFile('https://github.com/serbinskis/microsoft-office-2016/raw/refs/heads/master/Office2016/Office2016.zip.003', 'Office2016.zip.003')"
powershell -Command "(New-Object Net.WebClient).DownloadFile('https://github.com/serbinskis/microsoft-office-2016/raw/refs/heads/master/Office2016/Office2016.zip.004', 'Office2016.zip.004')"
powershell -Command "(New-Object Net.WebClient).DownloadFile('https://github.com/serbinskis/microsoft-office-2016/raw/refs/heads/master/Office2016/Office2016.zip.005', 'Office2016.zip.005')"
powershell -Command "(New-Object Net.WebClient).DownloadFile('https://github.com/serbinskis/microsoft-office-2016/raw/refs/heads/master/Office2016/Office2016.zip.006', 'Office2016.zip.006')"
powershell -Command "(New-Object Net.WebClient).DownloadFile('https://github.com/serbinskis/microsoft-office-2016/raw/refs/heads/master/Office2016/Office2016.zip.007', 'Office2016.zip.007')"
powershell -Command "(New-Object Net.WebClient).DownloadFile('https://github.com/serbinskis/microsoft-office-2016/raw/refs/heads/master/Office2016/Office2016.zip.008', 'Office2016.zip.008')"
powershell -Command "(New-Object Net.WebClient).DownloadFile('https://github.com/serbinskis/microsoft-office-2016/raw/refs/heads/master/Office2016/Office2016.zip.009', 'Office2016.zip.009')"
powershell -Command "(New-Object Net.WebClient).DownloadFile('https://github.com/serbinskis/microsoft-office-2016/raw/refs/heads/master/Office2016/Office2016.zip.010', 'Office2016.zip.010')"
powershell -Command "(New-Object Net.WebClient).DownloadFile('https://github.com/serbinskis/microsoft-office-2016/raw/refs/heads/master/Office2016/Office2016.zip.011', 'Office2016.zip.011')"
powershell -Command "(New-Object Net.WebClient).DownloadFile('https://github.com/serbinskis/microsoft-office-2016/raw/refs/heads/master/Office2016/Office2016.zip.012', 'Office2016.zip.012')"
powershell -Command "(New-Object Net.WebClient).DownloadFile('https://github.com/serbinskis/microsoft-office-2016/raw/refs/heads/master/Office2016/Office2016.zip.013', 'Office2016.zip.013')"
powershell -Command "(New-Object Net.WebClient).DownloadFile('https://github.com/serbinskis/microsoft-office-2016/raw/refs/heads/master/Office2016/Office2016.zip.014', 'Office2016.zip.014')"
powershell -Command "(New-Object Net.WebClient).DownloadFile('https://github.com/serbinskis/microsoft-office-2016/raw/refs/heads/master/Office2016/Office2016.zip.015', 'Office2016.zip.015')"
powershell -Command "(New-Object Net.WebClient).DownloadFile('https://github.com/serbinskis/microsoft-office-2016/raw/refs/heads/master/Office2016/Office2016.zip.016', 'Office2016.zip.016')"
powershell -Command "(New-Object Net.WebClient).DownloadFile('https://github.com/serbinskis/microsoft-office-2016/raw/refs/heads/master/Office2016/Office2016.zip.017', 'Office2016.zip.017')"
powershell -Command "(New-Object Net.WebClient).DownloadFile('https://github.com/serbinskis/microsoft-office-2016/raw/refs/heads/master/Office2016/Office2016.zip.018', 'Office2016.zip.018')"
powershell -Command "(New-Object Net.WebClient).DownloadFile('https://github.com/serbinskis/microsoft-office-2016/raw/refs/heads/master/Office2016/Office2016.zip.019', 'Office2016.zip.019')"

title Extracting - Microsoft Office 2016
powershell -Command "Invoke-WebRequest -Uri 'https://www.7-zip.org/a/7za920.zip' -OutFile '7za920.zip'"
powershell -Command "Expand-Archive -Path '7za920.zip' -DestinationPath '.'"
7za x Office2016.zip.001 >nul
del /s /q Office2016.zip.* >nul

title Installing - Microsoft Office 2016
if exist Configure.xml del Configure.xml

(echo | set /p var="<Configuration>") >> Configure.xml
echo. >> Configure.xml
(echo | set /p var="<Add SourcePath="%CD%" OfficeClientEdition="32">") >> Configure.xml
echo. >> Configure.xml
(echo | set /p var="<Product ID="WordRetail">") >> Configure.xml
echo. >> Configure.xml
(echo | set /p var="<Language ID="en-US" />") >> Configure.xml
echo. >> Configure.xml
(echo | set /p var="</Product>") >> Configure.xml
echo. >> Configure.xml
(echo | set /p var="<Product ID="ExcelRetail">") >> Configure.xml
echo. >> Configure.xml
(echo | set /p var="<Language ID="en-US" />") >> Configure.xml
echo. >> Configure.xml
(echo | set /p var="</Product>") >> Configure.xml
echo. >> Configure.xml
(echo | set /p var="<Product ID="PowerPointRetail">") >> Configure.xml
echo. >> Configure.xml
(echo | set /p var="<Language ID="en-US" />") >> Configure.xml
echo. >> Configure.xml
(echo | set /p var="</Product>") >> Configure.xml
echo. >> Configure.xml
(echo | set /p var="</Add>") >> Configure.xml
echo. >> Configure.xml
(echo | set /p var="<Display Level="Full" AcceptEULA="TRUE" />") >> Configure.xml
echo. >> Configure.xml
(echo | set /p var="<Updates Enabled="TRUE" Channel="Current" />") >> Configure.xml
echo. >> Configure.xml
(echo | set /p var="</Configuration>") >> Configure.xml
echo. >> Configure.xml
(echo | set /p var="<!--  OInstallSetting Branch=0  -->") >> Configure.xml
echo. >> Configure.xml
(echo | set /p var="<!--  OInstallSetting Office=0  -->") >> Configure.xml
echo. >> Configure.xml
(echo | set /p var="<!--  OInstallSetting Arch=0  -->") >> Configure.xml
echo. >> Configure.xml
(echo | set /p var="<!--  OInstallSetting Lang=0  --> ") >> Configure.xml
echo. >> Configure.xml

setup.exe /configure Configure.xml
taskkill /F /IM OfficeC2RClient.exe >nul 2>nul
taskkill /F /IM OneDriveSetup.exe >nul 2>nul
del Configure.xml >nul

timeout /t 2 /nobreak >nul
rd /s /q "C:\Program Files (x86)\Microsoft Office\root\Integration\Addons"
rd /s /q "C:\ProgramData\Microsoft OneDrive"
rd /s /q "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office Tools"
schtasks.exe /Change /TN "Microsoft\Office\Office Automatic Updates 2.0" /Disable >nul 2>nul
printui.exe /dl /n "Send To OneNote 2016" >nul 2>nul

reg delete HKEY_CLASSES_ROOT\.docx\Word.Document.12\ShellNew /f >nul 2>nul
reg delete HKEY_CLASSES_ROOT\.pptx\PowerPoint.Show.12\ShellNew /f >nul 2>nul
reg delete HKEY_CLASSES_ROOT\.xlsx\Excel.Sheet.12\ShellNew /f >nul 2>nul

cd "%TEMP%"
rd /s /q "%TEMP%\Office2016"
