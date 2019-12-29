
SET BUILDCONFIG=Release

FOR /F %%I IN ("%0") DO SET CURRENTDIR=%%~dpI

msbuild %CURRENTDIR%WindowsPE.csproj /p:Configuration=%BUILDCONFIG%;DefineConstants="TRACE" /t:Rebuild

robocopy %CURRENTDIR%bin\%BUILDCONFIG% %CURRENTDIR%nuget\root\lib\net40
robocopy %CURRENTDIR% %CURRENTDIR%nuget\root WindowsPE.nuspec

nuget pack %CURRENTDIR%\nuget\root\WindowsPE.nuspec -OutputDirectory %CURRENTDIR%nuget_output
if %ERRORLEVEL% GTR 0 goto BuildError

for /f %%i in ('powershell.exe -file %CURRENTDIR%..\..\..\getNugetVersion.ps1 %CURRENTDIR%nuget\root\WindowsPE.nuspec') do set NUGETVERSION=%%i
echo %NUGETVERSION%

for /f %%i in ('powershell.exe -file %CURRENTDIR%..\..\..\readFile.ps1 d:\settings\nuget_key.txt') do set NUGETKEY=%%i
echo %NUGETKEY%

nuget push %CURRENTDIR%nuget_output\WindowsPE.%NUGETVERSION%.nupkg %NUGETKEY% -src https://www.nuget.org/api/v2/package
goto EndOfBuild

:BuildError
Echo Failed to build


:EndOfBuild