@echo off

set MSVS_Cmd=H:\VisualStudio\Common7\Tools\VsDevCmd.bat

if not exist "%MSVS_Cmd%" (echo ERROR: Cannot found file: %MSVS_Cmd% & pause >NUL & goto :eof)

call "%MSVS_Cmd%"

cd /d "%~dp0"

del IRegexp.tlb 2>NUL

midl.exe /mktyplib203 IRegexp.odl

pause

