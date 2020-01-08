
FOR /F %%I IN ("%0") DO SET CURRENTDIR=%%~dpI

SET BUILDCONFIG=Release

REM ========= InstallDriver =================
msbuild "InstallDriver.vcxproj" /p:Platform=Win32;Configuration=%BUILDCONFIG% /p:TargetName=InstallDriver32 
if ERRORLEVEL 1 pause

msbuild "InstallDriver.vcxproj" /p:Platform=x64;Configuration=%BUILDCONFIG% /p:TargetName=InstallDriver64
if ERRORLEVEL 1 pause

robocopy %CURRENTDIR%x64\%BUILDCONFIG%  %CURRENTDIR%%BUILDCONFIG%  InstallDriver64.*


:EndOfBuild