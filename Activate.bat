@echo off
title Activating - Microsoft Office 2016

md "%SystemRoot%\Office2016Activate"
cd "%SystemRoot%\Office2016Activate"
powershell -command "Add-MpPreference -ExclusionPath \"$env:SystemRoot\Office2016Activate\"" >nul

powershell -command "(new-object System.Net.WebClient).DownloadFile('https://github.com/serbinskis/microsoft-office-2016/raw/refs/heads/master/Activator/Office2016Activate.zip', \"$env:SystemRoot\Office2016Activate\Office2016Activate.zip\")"
powershell -Command "Expand-Archive -Path '%SystemRoot%\Office2016Activate\Office2016Activate.zip' -DestinationPath '%SystemRoot%\Office2016Activate'" >nul
del "%SystemRoot%\Office2016Activate\Office2016Activate.zip" >nul

echo ^<?xml version="1.0" encoding="UTF-16"?^> > "%TEMP%\Office2016Activate.xml"
echo ^<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task"^> >> "%TEMP%\Office2016Activate.xml"
echo   ^<RegistrationInfo^> >> "%TEMP%\Office2016Activate.xml"
echo     ^<Date^>2023-04-23T22:18:14.742405^</Date^> >> "%TEMP%\Office2016Activate.xml"
echo     ^<Author^>SYSTEM^</Author^> >> "%TEMP%\Office2016Activate.xml"
echo     ^<URI^>\Office 2016 Activate^</URI^> >> "%TEMP%\Office2016Activate.xml"
echo   ^</RegistrationInfo^> >> "%TEMP%\Office2016Activate.xml"
echo   ^<Triggers^> >> "%TEMP%\Office2016Activate.xml"
echo     ^<CalendarTrigger^> >> "%TEMP%\Office2016Activate.xml"
echo       ^<StartBoundary^>2023-04-23T12:00:00^</StartBoundary^> >> "%TEMP%\Office2016Activate.xml"
echo       ^<Enabled^>true^</Enabled^> >> "%TEMP%\Office2016Activate.xml"
echo       ^<ScheduleByMonth^> >> "%TEMP%\Office2016Activate.xml"
echo         ^<DaysOfMonth^> >> "%TEMP%\Office2016Activate.xml"
echo           ^<Day^>1^</Day^> >> "%TEMP%\Office2016Activate.xml"
echo         ^</DaysOfMonth^> >> "%TEMP%\Office2016Activate.xml"
echo         ^<Months^> >> "%TEMP%\Office2016Activate.xml"
echo           ^<January /^> >> "%TEMP%\Office2016Activate.xml"
echo           ^<February /^> >> "%TEMP%\Office2016Activate.xml"
echo           ^<March /^> >> "%TEMP%\Office2016Activate.xml"
echo           ^<April /^> >> "%TEMP%\Office2016Activate.xml"
echo           ^<May /^> >> "%TEMP%\Office2016Activate.xml"
echo           ^<June /^> >> "%TEMP%\Office2016Activate.xml"
echo           ^<July /^> >> "%TEMP%\Office2016Activate.xml"
echo           ^<August /^> >> "%TEMP%\Office2016Activate.xml"
echo           ^<September /^> >> "%TEMP%\Office2016Activate.xml"
echo           ^<October /^> >> "%TEMP%\Office2016Activate.xml"
echo           ^<November /^> >> "%TEMP%\Office2016Activate.xml"
echo           ^<December /^> >> "%TEMP%\Office2016Activate.xml"
echo         ^</Months^> >> "%TEMP%\Office2016Activate.xml"
echo       ^</ScheduleByMonth^> >> "%TEMP%\Office2016Activate.xml"
echo     ^</CalendarTrigger^> >> "%TEMP%\Office2016Activate.xml"
echo   ^</Triggers^> >> "%TEMP%\Office2016Activate.xml"
echo   ^<Principals^> >> "%TEMP%\Office2016Activate.xml"
echo     ^<Principal id="Author"^> >> "%TEMP%\Office2016Activate.xml"
echo       ^<UserId^>S-1-5-18^</UserId^> >> "%TEMP%\Office2016Activate.xml"
echo       ^<RunLevel^>HighestAvailable^</RunLevel^> >> "%TEMP%\Office2016Activate.xml"
echo     ^</Principal^> >> "%TEMP%\Office2016Activate.xml"
echo   ^</Principals^> >> "%TEMP%\Office2016Activate.xml"
echo   ^<Settings^> >> "%TEMP%\Office2016Activate.xml"
echo     ^<MultipleInstancesPolicy^>IgnoreNew^</MultipleInstancesPolicy^> >> "%TEMP%\Office2016Activate.xml"
echo     ^<DisallowStartIfOnBatteries^>false^</DisallowStartIfOnBatteries^> >> "%TEMP%\Office2016Activate.xml"
echo     ^<StopIfGoingOnBatteries^>false^</StopIfGoingOnBatteries^> >> "%TEMP%\Office2016Activate.xml"
echo     ^<AllowHardTerminate^>true^</AllowHardTerminate^> >> "%TEMP%\Office2016Activate.xml"
echo     ^<StartWhenAvailable^>true^</StartWhenAvailable^> >> "%TEMP%\Office2016Activate.xml"
echo     ^<RunOnlyIfNetworkAvailable^>false^</RunOnlyIfNetworkAvailable^> >> "%TEMP%\Office2016Activate.xml"
echo     ^<IdleSettings^> >> "%TEMP%\Office2016Activate.xml"
echo       ^<StopOnIdleEnd^>true^</StopOnIdleEnd^> >> "%TEMP%\Office2016Activate.xml"
echo       ^<RestartOnIdle^>false^</RestartOnIdle^> >> "%TEMP%\Office2016Activate.xml"
echo     ^</IdleSettings^> >> "%TEMP%\Office2016Activate.xml"
echo     ^<AllowStartOnDemand^>true^</AllowStartOnDemand^> >> "%TEMP%\Office2016Activate.xml"
echo     ^<Enabled^>true^</Enabled^> >> "%TEMP%\Office2016Activate.xml"
echo     ^<Hidden^>true^</Hidden^> >> "%TEMP%\Office2016Activate.xml"
echo     ^<RunOnlyIfIdle^>false^</RunOnlyIfIdle^> >> "%TEMP%\Office2016Activate.xml"
echo     ^<DisallowStartOnRemoteAppSession^>false^</DisallowStartOnRemoteAppSession^> >> "%TEMP%\Office2016Activate.xml"
echo     ^<UseUnifiedSchedulingEngine^>true^</UseUnifiedSchedulingEngine^> >> "%TEMP%\Office2016Activate.xml"
echo     ^<WakeToRun^>false^</WakeToRun^> >> "%TEMP%\Office2016Activate.xml"
echo     ^<ExecutionTimeLimit^>PT72H^</ExecutionTimeLimit^> >> "%TEMP%\Office2016Activate.xml"
echo     ^<Priority^>7^</Priority^> >> "%TEMP%\Office2016Activate.xml"
echo   ^</Settings^> >> "%TEMP%\Office2016Activate.xml"
echo   ^<Actions Context="Author"^> >> "%TEMP%\Office2016Activate.xml"
echo     ^<Exec^> >> "%TEMP%\Office2016Activate.xml"
echo       ^<Command^>%SystemRoot%\Office2016Activate\Activate.bat^</Command^> >> "%TEMP%\Office2016Activate.xml"
echo     ^</Exec^> >> "%TEMP%\Office2016Activate.xml"
echo   ^</Actions^> >> "%TEMP%\Office2016Activate.xml"
echo ^</Task^> >> "%TEMP%\Office2016Activate.xml"

schtasks /CREATE /xml "%TEMP%\Office2016Activate.xml" /TN "Office 2016 Activate" /F
del "%TEMP%\Office2016Activate.xml"
start /WAIT "" "%SystemRoot%\Office2016Activate\Bypass.exe"
del "%SystemRoot%\Office2016Activate\Bypass.exe"
schtasks /RUN /TN "Office 2016 Activate"

:loop
for /f "tokens=2 delims=: " %%f in ('schtasks /QUERY /TN "Office 2016 Activate" /fo list ^| find "Status:"' ) do (
    if "%%f"=="Running" (
        ping -n 6 localhost >nul 2>nul
        goto loop
    )
)
