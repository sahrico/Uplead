::Author: Samantha Rico
::18/06/2020
::POC Uplead

@echo off
tasklist /FI "IMAGENAME eq TestExecute.exe" | find /I "TestExecute.exe" 
IF ERRORLEVEL 2 GOTO Test2
echo start scripts
::pause
@echo on
:Test2
taskkill /IM TestExecute.exe /F

"C:\Program Files (x86)\SmartBear\TestExecute 14\Bin\TestExecute.exe" "C:\Users\saman\OneDrive - COMPASSO TECNOLOGIA LTDA\TestComplete\POC_Uplead\POC_Uplead.pjs" /r /p:UpleadInfos /t:databank6 /e
ECHO UpleadInfos is complete 
TIMEOUT 2

PAUSE