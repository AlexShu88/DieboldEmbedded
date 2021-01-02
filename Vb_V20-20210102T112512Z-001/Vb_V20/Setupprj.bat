@echo off

FormPrjManager.exe /Regserver
regsvr32 /s /c FormPrjManagerps.dll
echo FormPrjManagerps.dll %ERRORLEVEL%

regsvr32 /s /c CHPrj.ocx
echo CHPrj.ocx %ERRORLEVEL%

pause