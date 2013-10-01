@echo off

::::run npp installer first
::for %%i in (*.exe) do start /wait %%i

set programfiles=C:\Program Files\Notepad++
if exist "%appdata%\Notepad++" (
	set defaultloc=%appdata%\Notepad++
) else (
	set defaultloc=%programfiles%
)

copy *.dll "%programfiles%\plugins"
copy *.xml "%defaultloc%"

cd functionlist

copy FunctionList.dll "%programfiles%\plugins"
copy "Gmod Lua.bmp" "%programfiles%\plugins\config"
copy "C++.flb" "%programfiles%\plugins\config"
copy FunctionListRules.xml "%defaultloc%\plugins\config"

pause