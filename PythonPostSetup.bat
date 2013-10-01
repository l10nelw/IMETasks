@echo off
if defined PythonPath goto FINISH

echo Setting up environment variables...
set key="HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
set setpythonpath=c:\python32;c:\python32\scripts;c:\imetasks
for /f "skip=4 tokens=2*" %%i in ('reg query %key% /v Path') do set newpath=%%j;%%PythonPath%%
reg add %key% /v PythonPath /d "%setpythonpath%" /f
reg add %key% /v Path /t REG_EXPAND_SZ /d "%newpath%" /f

:FINISH
echo Done.
pause