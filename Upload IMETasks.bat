@echo off
set dest="P:\Media Content\Internet Monitoring\Internet Monitoring Service Team\IM Engineers\Programs\IMETasks"
echo Copying files to %dest%...
echo.
for %%i in (imetasks.bat imetasks.ini listjob.py) do (
	echo %%i
	copy %%i %dest%
)
echo Done.
echo.
pause