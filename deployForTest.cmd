Set /P answr=deploy (r)elease? 
set source=bin\Debug
If "%answr%"=="r" (
set source=bin\Release
)
copy /Y %source%\Raddin-AddIn-packed.xll "%appdata%\Microsoft\AddIns"
copy /Y %source%\Raddin.pdb "%appdata%\Microsoft\AddIns"
pause

