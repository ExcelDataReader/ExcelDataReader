@echo off

REM ************Compact Framework. DEBUG to Bin\Release\CF\Excel.dll************
REM **********************************************************************
SET CSC=C:\WINDOWS\Microsoft.NET\Framework\v2.0.50727\csc.exe /nologo

SET SC_PATH=%cd%

REM **CLEAN

RD /S /Q %SC_PATH%\Bin\Debug\CF\
MKDIR %SC_PATH%\Bin\Debug\CF\

if "%NETCF_PATH%" == "" (
**set NETCF_PATH=C:\Program Files\Microsoft.NET\SDK\v2.0\CompactFramework\WindowsCE)
  set NETCF_PATH=C:\Program Files\Microsoft Visual Studio 8\SmartDevices\SDK\CompactFramework\2.0\v2.0\WindowsCE

if DEFINED REF ( set REF= )

set REF=%REF% "/r:%NETCF_PATH%\MsCorlib.dll"
set REF=%REF% "/r:%NETCF_PATH%\System.dll"
set REF=%REF% "/r:%NETCF_PATH%\System.Data.dll"
set REF=%REF% "/r:%NETCF_PATH%\System.Xml.dll"
set REF=%REF% "/r:%SC_PATH%\Binaries\netcf-20\ICSharpCode.SharpZipLib.dll"

%CSC% /nologo /debug -nostdlib -noconfig  /out:%SC_PATH%\Bin\Debug\CF\Excel.dll /target:library %REF% /recurse:*.cs

REM **********************************************************************