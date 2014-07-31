@echo off

REM ************Compact Framework. RELEASE to Bin\Release\CF\Excel.dll************
REM **********************************************************************
SET CSC=C:\WINDOWS\Microsoft.NET\Framework\v2.0.50727\csc.exe /nologo

SET SC_PATH=%cd%

REM **CLEAN

RD /S /Q %SC_PATH%\Bin\Release\CF\
MKDIR %SC_PATH%\Bin\Release\CF\

set NETCF_PATH=C:\Program Files\Microsoft.NET\SDK\CompactFramework\v2.0\WindowsCE


if DEFINED REF ( set REF= )

set REF=%REF% "/r:%NETCF_PATH%\MsCorlib.dll"
set REF=%REF% "/r:%NETCF_PATH%\System.dll"
set REF=%REF% "/r:%NETCF_PATH%\System.Data.dll"
set REF=%REF% "/r:%NETCF_PATH%\System.Xml.dll"
set REF=%REF% /r:"%SC_PATH%\..\Lib\cf\ICSharpCode.SharpZipLib.dll"

@echo on
%CSC% /define:CF_RELEASE /nologo -nostdlib -noconfig /o /out:"%SC_PATH%\Bin\Release\CF\Excel.dll" /target:library %REF% /recurse:*.cs