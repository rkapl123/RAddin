Set /P answr=deploy (r)elease (empty for debug)? 
set source=bin\Debug
If "%answr%"=="r" (
set source=bin\Release
copy /Y %source%\Raddin-AddIn64-packed.xll Distribution\Raddin64.xll
copy /Y %source%\Raddin.dll.config Distribution
copy /Y %source%\Raddin-AddIn-packed.xll Distribution\Raddin32.xll
)
if exist "C:\Program Files\Microsoft Office\root\" (
	echo 64bit office
	copy /Y %source%\Raddin-AddIn64-packed.xll "%appdata%\Microsoft\AddIns\Raddin.xll"
	copy /Y %source%\Raddin.pdb "%appdata%\Microsoft\AddIns"
	copy /Y %source%\Raddin.dll.config "%appdata%\Microsoft\AddIns\Raddin.xll.config"
) else (
	echo 32bit office
	copy /Y %source%\Raddin-AddIn-packed.xll "%appdata%\Microsoft\AddIns\Raddin.xll"
	copy /Y %source%\Raddin.pdb "%appdata%\Microsoft\AddIns"
	copy /Y %source%\Raddin.dll.config "%appdata%\Microsoft\AddIns\Raddin.xll.config"
)
pause
