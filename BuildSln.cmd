@ECHO OFF
SETLOCAL
PATH=%ProgramFiles(x86)%\Microsoft Visual Studio\2017\Professional\MSBuild\15.0\Bin;%programfiles(x86)%\Microsoft Visual Studio\2017\Enterprise\MSBuild\15.0\Bin;%ProgramFiles(x86)%\MSBuild\14.0\Bin;%PATH%

TITLE Build solution in release for any cpu

ECHO Delete **\bin
FOR /F "tokens=*" %%G IN ('DIR /B /AD /S bin') DO RMDIR /S /Q "%%G"

ECHO Delete **\obj
FOR /F "tokens=*" %%G IN ('DIR /B /AD /S obj') DO RMDIR /S /Q "%%G"

ECHO Delete packages
RMDIR packages /S /Q

ECHO Restore nuget packages
.nuget\nuget.exe restore OnCallOutlookEntry.sln

ECHO Build Client.sln
MSBuild.exe OnCallOutlookEntry.sln /m /fl /nologo /t:Clean,ReBuild /clp:ShowTimestamp;ErrorsOnly;Summary /p:WarningLevel=1;SkipInvalidConfigurations=true;Configuration="Release";Platform="Any CPU"

ENDLOCAL
PAUSE
