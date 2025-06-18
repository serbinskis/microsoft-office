@echo off
title Activating - Microsoft Office

md "%SystemRoot%\OfficeActivate"
powershell -command "Add-MpPreference -ExclusionPath \"$env:SystemRoot\OfficeActivate\"" >nul
powershell -command "(new-object System.Net.WebClient).DownloadFile('https://github.com/serbinskis/microsoft-office/raw/refs/heads/master/Activator/OfficeActivate.zip', \"$env:SystemRoot\OfficeActivate\OfficeActivate.zip\")"
powershell -command "Expand-Archive -Path '%SystemRoot%\OfficeActivate\OfficeActivate.zip' -DestinationPath '%SystemRoot%\OfficeActivate'" >nul
del "%SystemRoot%\OfficeActivate\OfficeActivate.zip" >nul

echo ^<?xml version="1.0" encoding="UTF-16"?^> > "%TEMP%\OfficeActivate.xml"
echo ^<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task"^> >> "%TEMP%\OfficeActivate.xml"
echo   ^<RegistrationInfo^> >> "%TEMP%\OfficeActivate.xml"
echo     ^<Date^>2023-04-23T22:18:14.742405^</Date^> >> "%TEMP%\OfficeActivate.xml"
echo     ^<Author^>SYSTEM^</Author^> >> "%TEMP%\OfficeActivate.xml"
echo     ^<URI^>\Office Activate^</URI^> >> "%TEMP%\OfficeActivate.xml"
echo   ^</RegistrationInfo^> >> "%TEMP%\OfficeActivate.xml"
echo   ^<Triggers^> >> "%TEMP%\OfficeActivate.xml"
echo     ^<CalendarTrigger^> >> "%TEMP%\OfficeActivate.xml"
echo       ^<StartBoundary^>2023-04-23T12:00:00^</StartBoundary^> >> "%TEMP%\OfficeActivate.xml"
echo       ^<Enabled^>true^</Enabled^> >> "%TEMP%\OfficeActivate.xml"
echo       ^<ScheduleByMonth^> >> "%TEMP%\OfficeActivate.xml"
echo         ^<DaysOfMonth^> >> "%TEMP%\OfficeActivate.xml"
echo           ^<Day^>1^</Day^> >> "%TEMP%\OfficeActivate.xml"
echo         ^</DaysOfMonth^> >> "%TEMP%\OfficeActivate.xml"
echo         ^<Months^> >> "%TEMP%\OfficeActivate.xml"
echo           ^<January /^> >> "%TEMP%\OfficeActivate.xml"
echo           ^<February /^> >> "%TEMP%\OfficeActivate.xml"
echo           ^<March /^> >> "%TEMP%\OfficeActivate.xml"
echo           ^<April /^> >> "%TEMP%\OfficeActivate.xml"
echo           ^<May /^> >> "%TEMP%\OfficeActivate.xml"
echo           ^<June /^> >> "%TEMP%\OfficeActivate.xml"
echo           ^<July /^> >> "%TEMP%\OfficeActivate.xml"
echo           ^<August /^> >> "%TEMP%\OfficeActivate.xml"
echo           ^<September /^> >> "%TEMP%\OfficeActivate.xml"
echo           ^<October /^> >> "%TEMP%\OfficeActivate.xml"
echo           ^<November /^> >> "%TEMP%\OfficeActivate.xml"
echo           ^<December /^> >> "%TEMP%\OfficeActivate.xml"
echo         ^</Months^> >> "%TEMP%\OfficeActivate.xml"
echo       ^</ScheduleByMonth^> >> "%TEMP%\OfficeActivate.xml"
echo     ^</CalendarTrigger^> >> "%TEMP%\OfficeActivate.xml"
echo   ^</Triggers^> >> "%TEMP%\OfficeActivate.xml"
echo   ^<Principals^> >> "%TEMP%\OfficeActivate.xml"
echo     ^<Principal id="Author"^> >> "%TEMP%\OfficeActivate.xml"
echo       ^<UserId^>S-1-5-18^</UserId^> >> "%TEMP%\OfficeActivate.xml"
echo       ^<RunLevel^>HighestAvailable^</RunLevel^> >> "%TEMP%\OfficeActivate.xml"
echo     ^</Principal^> >> "%TEMP%\OfficeActivate.xml"
echo   ^</Principals^> >> "%TEMP%\OfficeActivate.xml"
echo   ^<Settings^> >> "%TEMP%\OfficeActivate.xml"
echo     ^<MultipleInstancesPolicy^>IgnoreNew^</MultipleInstancesPolicy^> >> "%TEMP%\OfficeActivate.xml"
echo     ^<DisallowStartIfOnBatteries^>false^</DisallowStartIfOnBatteries^> >> "%TEMP%\OfficeActivate.xml"
echo     ^<StopIfGoingOnBatteries^>false^</StopIfGoingOnBatteries^> >> "%TEMP%\OfficeActivate.xml"
echo     ^<AllowHardTerminate^>true^</AllowHardTerminate^> >> "%TEMP%\OfficeActivate.xml"
echo     ^<StartWhenAvailable^>true^</StartWhenAvailable^> >> "%TEMP%\OfficeActivate.xml"
echo     ^<RunOnlyIfNetworkAvailable^>false^</RunOnlyIfNetworkAvailable^> >> "%TEMP%\OfficeActivate.xml"
echo     ^<IdleSettings^> >> "%TEMP%\OfficeActivate.xml"
echo       ^<StopOnIdleEnd^>true^</StopOnIdleEnd^> >> "%TEMP%\OfficeActivate.xml"
echo       ^<RestartOnIdle^>false^</RestartOnIdle^> >> "%TEMP%\OfficeActivate.xml"
echo     ^</IdleSettings^> >> "%TEMP%\OfficeActivate.xml"
echo     ^<AllowStartOnDemand^>true^</AllowStartOnDemand^> >> "%TEMP%\OfficeActivate.xml"
echo     ^<Enabled^>true^</Enabled^> >> "%TEMP%\OfficeActivate.xml"
echo     ^<Hidden^>true^</Hidden^> >> "%TEMP%\OfficeActivate.xml"
echo     ^<RunOnlyIfIdle^>false^</RunOnlyIfIdle^> >> "%TEMP%\OfficeActivate.xml"
echo     ^<DisallowStartOnRemoteAppSession^>false^</DisallowStartOnRemoteAppSession^> >> "%TEMP%\OfficeActivate.xml"
echo     ^<UseUnifiedSchedulingEngine^>true^</UseUnifiedSchedulingEngine^> >> "%TEMP%\OfficeActivate.xml"
echo     ^<WakeToRun^>false^</WakeToRun^> >> "%TEMP%\OfficeActivate.xml"
echo     ^<ExecutionTimeLimit^>PT72H^</ExecutionTimeLimit^> >> "%TEMP%\OfficeActivate.xml"
echo     ^<Priority^>7^</Priority^> >> "%TEMP%\OfficeActivate.xml"
echo   ^</Settings^> >> "%TEMP%\OfficeActivate.xml"
echo   ^<Actions Context="Author"^> >> "%TEMP%\OfficeActivate.xml"
echo     ^<Exec^> >> "%TEMP%\OfficeActivate.xml"
echo       ^<Command^>%SystemRoot%\OfficeActivate\Activate.bat^</Command^> >> "%TEMP%\OfficeActivate.xml"
echo     ^</Exec^> >> "%TEMP%\OfficeActivate.xml"
echo   ^</Actions^> >> "%TEMP%\OfficeActivate.xml"
echo ^</Task^> >> "%TEMP%\OfficeActivate.xml"

schtasks /create /xml "%TEMP%\OfficeActivate.xml" /tn "Office Activate" /f >nul 2>nul
del "%TEMP%\OfficeActivate.xml"
start /WAIT "" "%SystemRoot%\OfficeActivate\Bypass.exe"
del "%SystemRoot%\OfficeActivate\Bypass.exe"
schtasks /run /tn "Office Activate" >nul 2>nul

:loop
for /f "tokens=2 delims=: " %%f in ('schtasks /query /tb "Office Activate" /fo list ^| find "Status:"' ) do (
    if "%%f"=="Running" (
        ping -n 6 localhost >nul 2>nul
        goto loop
    )
)
