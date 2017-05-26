@echo OFF
setlocal ENABLEEXTENSIONS
FOR /F "usebackq skip=2 tokens=1 delims=" %%A IN (`REG QUERY HKLM\Software\R-core\R /v InstallPath 2^>nul`) DO (
    set result=%%A
)
SET rpath=%result:*REG_SZ    =%
"%rpath%\bin\Rscript.exe" %~dp0\TestScript.R
pause