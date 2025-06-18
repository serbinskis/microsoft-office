@echo off
title Installing - Microsoft Office

REM Download Office Click To Run
powershell -command "New-Item -ItemType Directory -Force -Path '%SystemDrive%\Program Files\Common Files\Microsoft Shared\ClickToRun'" >nul
powershell -command "(New-Object Net.WebClient).DownloadFile('http://officecdn.microsoft.com/pr/7983bac0-e531-40cf-be00-fd24fe66619c/Office/Data/16.0.17932.20396/i640.cab', '%TEMP%\office_i640.cab')"
powershell -command "(New-Object Net.WebClient).DownloadFile('http://officecdn.microsoft.com/pr/7983bac0-e531-40cf-be00-fd24fe66619c/Office/Data/16.0.17932.20396/i641033.cab', '%TEMP%\office_i641033.cab')"
powershell -command "(New-Object Net.WebClient).DownloadFile('http://officecdn.microsoft.com/pr/7983bac0-e531-40cf-be00-fd24fe66619c/Office/Data/16.0.17932.20396/i641062.cab', '%TEMP%\office_i641062.cab')"
powershell -command "(New-Object Net.WebClient).DownloadFile('http://officecdn.microsoft.com/pr/7983bac0-e531-40cf-be00-fd24fe66619c/Office/Data/16.0.17932.20396/i641049.cab', '%TEMP%\office_i641049.cab')"

REM Extract Office Click To Run
expand "%TEMP%\office_i640.cab" -F:* "%SystemDrive%\Program Files\Common Files\Microsoft Shared\ClickToRun" >nul
expand "%TEMP%\office_i641033.cab" -F:* "%SystemDrive%\Program Files\Common Files\Microsoft Shared\ClickToRun" >nul
expand "%TEMP%\office_i641062.cab" -F:* "%SystemDrive%\Program Files\Common Files\Microsoft Shared\ClickToRun" >nul
expand "%TEMP%\office_i641049.cab" -F:* "%SystemDrive%\Program Files\Common Files\Microsoft Shared\ClickToRun" >nul
del "%TEMP%\office_*.cab" >nul

REM Start Office Installation
start /WAIT /MIN "" "%SystemDrive%\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" deliverymechanism=7983bac0-e531-40cf-be00-fd24fe66619c platform=x64 productreleaseid=none culture=en-us defaultplatform=False lcid=1033 b= storeid= forceupgrade=True piniconstotaskbar=False pidkeys=MQ84N-7VYDM-FXV7C-6K7CC-VFW9J,F4DYN-89BP2-WQTWJ-GR8YC-CKGJG forceappshutdown=True autoactivate=1 productstoadd="Word2024Volume.16_en-us_ru-ru_lv-lv_x-none|Excel2024Volume.16_en-us_ru-ru_lv-lv_x-none|PowerPoint2024Volume.16_en-us_ru-ru_lv-lv_x-none|ProofingTools.16_en-us_ru-ru_lv-lv_x-none" scenario=unknown updatesenabled.16=False acceptalleulas.16=True cdnbaseurl.16=http://officecdn.microsoft.com/pr/7983bac0-e531-40cf-be00-fd24fe66619c version.16=16.0.17932.20396 mediatype.16=CDN baseurl.16=http://officecdn.microsoft.com/pr/7983bac0-e531-40cf-be00-fd24fe66619c sourcetype.16=CDN displaylevel=True uninstallpreviousversion=True Word2024Volume.excludedapps.16=onedrive Excel2024Volume.excludedapps.16=onedrive PowerPoint2024Volume.excludedapps.16=onedrive ProofingTools.excludedapps.16=onedrive

REM Remove Unnecessary Start Menu Shortcuts
rd /s /q "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office Tools"

REM Close Installation Finish Popup
timeout /t 3 /nobreak >nul & taskkill /f /im OfficeC2RClient.exe >nul 2>nul

REM Disable Any Scheduled Tasks For Office
powershell -command "Get-ScheduledTask -TaskPath '\Microsoft\Office\' | Disable-ScheduledTask" >nul

REM Remove Context Menu -> New -> File Types
reg delete HKEY_CLASSES_ROOT\.docx\Word.Document.12\ShellNew /f >nul 2>nul
reg delete HKEY_CLASSES_ROOT\.pptx\PowerPoint.Show.12\ShellNew /f >nul 2>nul
reg delete HKEY_CLASSES_ROOT\.xlsx\Excel.Sheet.12\ShellNew /f >nul 2>nul

REM echo Change Dark Theme, Accept EULA, Accept Data Sending, Disable PowerPoint Dialogs, Enable Word Ruler
REM reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\Common\Roaming" /v SchemaVersion /t REG_DWORD /d 9 /f >nul 2>nul
REM reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\Common\Roaming\Identities\Anonymous\Settings\1186\{00000000-0000-0000-0000-000000000000}\PendingChanges" /v Data /t REG_BINARY /d 04000000 /f >nul 2>nul
reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\Common" /v QMEnable /t REG_DWORD /d 1 /f >nul 2>nul
reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\Common\General" /v ShownFirstRunOptin /t REG_DWORD /d 1 /f >nul 2>nul
reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\Common\Licensing" /v EulasSetAccepted /t REG_SZ /d "8,0,49," /f >nul 2>nul
reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\Common\Privacy\SettingsStore\Anonymous" /v FRESettingsMigrated /t REG_DWORD /d 1 /f >nul 2>nul
reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\Common\Privacy\SettingsStore\Anonymous" /v RequiredDiagnosticDataNoticeVersion /t REG_DWORD /d 1 /f >nul 2>nul
reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\Common\Privacy\SettingsStore\Anonymous" /v OptionalDiagnosticDataConsentVersion /t REG_DWORD /d 1 /f >nul 2>nul
reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\Common\Privacy\SettingsStore\Anonymous" /v ConnectedExperiencesNoticeVersion /t REG_DWORD /d 1 /f >nul 2>nul
reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\Common\Privacy\SettingsStore\Anonymous" /v OptionalConnectedExperiencesNoticeVersion /t REG_DWORD /d 2 /f >nul 2>nul
reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\Common\PromoDialog" /v FluentWelcomeDialogDelay /t REG_DWORD /d 1 /f >nul 2>nul
reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\Common\PromoDialog" /v FluentWelcomeDialogShown /t REG_DWORD /d 1 /f >nul 2>nul
reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\Common\TeachingCallouts" /v PPT_RecordVideoPresentation_Callout /t REG_DWORD /d 2 /f >nul 2>nul
reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\Word\Options" /v Ruler /t REG_DWORD /d 1 /f >nul 2>nul