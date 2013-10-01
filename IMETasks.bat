@echo off
:: IMETasks.bat
:: Simple task launcher for Internet Monitoring Engineers
set contact=lionel.wong@isentia.com
:: =======================================================================================
:: === Please ensure the settings in IMETasks.ini are correct before running this app! ===
:: =======================================================================================

set version=1.14.0
:: 1.14.0	2013-08-30 11:04  c removed. support for double-digit spiders. Code readability improvements
:: 1.13.4	2013-07-16 14:16  time can have leading zero
:: 1.13.3	2013-04-18 17:20  Fixed bug to accommodate both local and Citrix Excel. Improved incomplete command handling
:: 1.13.2	2012-12-07 10:47  s9 added. ul renamed to ue. xl renamed to lx. Edited x message
:: 1.13.1	2012-08-16 10:20  lj fixed
:: 1.13.0	2012-08-15 14:10  lj added
:: 1.12.4	2012-07-10 19:15  fixed the x ask
:: 1.12.3	2012-06-26 20:16  f renamed to s. ss renamed to fs. retain spider.log when deleting temp files
:: 1.12.2	2012-06-13 11:40  xl added. HarvestSheet_Path renamed IMEngineers_Path
:: 1.12.1	2012-06-13 10:03  Fixed Notepad++ explicitness error in :UFF. bx requires min 4 backups before deleting. Cosmetics
:: 1.12.0	2012-06-12 17:25  sl added. ue renamed to ul. All plaintext files explicitly opened with Notepad++
:: 1.11.0	2012-06-08 17:51  p added. d renamed to da. bx now deletes oldest backups
:: 1.10.0	2012-05-15 16:10  c replaced with new command
:: 1.9.1	2012-05-03 10:44  fu fixed
:: 1.9.0	2012-04-25 10:40  ue added, refactored f
:: 1.8.1	2012-04-16 10:52  s renamed to f and fixed. fs renamed to ss
:: 1.8.0	2012-03-06 18:01  uf added. s improved. c hidden
:: 1.7.2	2012-03-01 11:06  s fixed and improved. Minor date/time display change
:: 1.7.1	2012-02-28 16:35  s now accommodates \urlfilefilter#\ paths
:: 1.7.0	2012-02-27 19:39  ss renamed to fs. s1 now works for any number
:: 1.6.3	2011-11-03 11:06  Automatic backup before opening harvest sheet
:: 1.6.2	2011-10-03 10:45  Restrict deletions in idx,logs folders to *.idx,*.log files
:: 1.6.1	2011-09-23 10:04  Serious backup bug fixed
:: 1.6.0	2011-09-21 11:10  Added IMETasks.ini. Happy birthday me!
:: 1.5.2	2011-09-15 09:54  Backup improved: create dir if not exist, fix dir location
:: 1.5.1	2011-08-30 15:16  Improved ss. Fixed empty prompt bug
:: 1.5.0	2011-08-03 12:51  Backup improved: call sub, local copy, lastsave.txt
:: 1.4.0	2011-08-01 17:17  Added Tools: ss search and c calc
:: 1.3.1	2011-05-20 14:35  Added HarvestSheetPath var
:: 1.3.0	2011-05-13 14:49  Added x: open harvest sheet
:: 1.2.0	2011-05-05 18:31  Added b: backup harvest sheet
:: 1.1.0	2011-05-04 13:29  Various improvements
:: 1.0.0	2011-04-06 10:29

:: load variables from imetasks.ini
for /f "eol=; tokens=1,2 delims==" %%i in (imetasks.ini) do set %%i=%%j

if not exist "%HarvestSheetBackup_Path%" md "%HarvestSheetBackup_Path%"

:: check for presence of local Excel
set excel=C:\Program Files\Microsoft Office\Office14
if exist "%excel%\excel.exe" ( path %path%;%excel% ) else set excel=no

set rundir=%cd%

color 1e
title IMETasks %version%
:MAIN
	echo.
	echo Internet Monitoring Engineer Task Suite %version%
	echo ---Local Testing----------------------------------------------------
	echo   h    Run HTTPConnector              u    Run URLFileFilter
	echo   oh   Open HTTPConnector.cfg         ou   Open FeedSettings.xml
	echo   fh   Open IDX folder                fu   Open URLFiles folder
	echo   sl   Open spider.log                hh   Open HTTPConnector Help
	echo ---Production-------------------------------------------------------
	echo   i    Launch IDOLConsole             ue#  Open \ErrorLogDir in S#
	echo   s#   Open FeedSettings.xml in S#    uf#  Open \UrlFiles in S#
	echo ---Data/Housekeeping------------------------------------------------
	echo   x    Open %HarvestSheet_File%
	echo   lx   Open last backed-up %HarvestSheet_File%
	echo   bx   Backup last saved %HarvestSheet_File%
	echo   lj   List most recently configured jobs in Excel-ready text format
	echo   dh   Delete HTTPConnector temp files (*.idx, *.log, *.status)
	echo   du   Delete URLFileFilter temp files (*.txt)
	echo   da   Delete all temp files
	echo ---Other------------------------------------------------------------
	echo   fs # Find string "#" in *.cfg       p    Update IMETasks
	echo   q    Quit
	echo.
	set x=
	set /p x=Enter command: 
	echo Command entered on %date% at %time%
	
	if "%x%"=="" goto :MAIN
	
	:: ---Local Testing----------------------------------------------------
	
	if "%x%"=="h" (
		echo Restarting HTTPConnector...
		net stop HTTPConnector > nul
		net start HTTPConnector
	)
	if "%x%"=="u" (
		echo ^| URLFileFilter should now be running, please wait for results.
		echo ^| Close or force close URLFileFilter at any time to continue
		echo ^| using IMETasks.
		cd /d %URLFileFilter_Path%
		URLFileFilter
	)
	if "%x%"=="oh" (
		echo Opening HTTPConnector.cfg...
		start notepad++ %HTTPConnector_Path%\HTTPConnector.cfg
	)
	if "%x%"=="ou" (
		echo Opening FeedSettings.xml...
		start notepad++ %URLFileFilter_Path%\FeedSettings.xml
	)
	if "%x%"=="hh" (
		echo Opening HTTPConnector Help...
		start %HTTPConnector_Path%\help\HTTPConnector\index.html
	)
	if "%x%"=="fh" (
		echo Opening IDX folder...
		start %IDX_Path%
	)
	if "%x%"=="fu" (
		echo Opening URLFiles folder...
		start %URLFiles_Path%
		goto :MAIN
	)
	if "%x%"=="sl" (
		echo Opening spider.log...
		cd /d %IDX_Path%
		if not exist spider.log rem. > spider.log
		start notepad++ spider.log
	)
	
	:: ---Production-------------------------------------------------------
	
	if "%x%"=="i" (
		echo Starting IDOLConsole...
		start %IdolConsole_Path%\idolconsole.exe
	)
	
	if "%x%"=="s" goto :incomplete
	if "%x:~0,1%"=="s" call :UFFPAths %x:~1% FeedSettings.xml notepad++

	if "%x%"=="ue" goto :incomplete
	if "%x:~0,2%"=="ue" call :UFFPAths %x:~2% ErrorLogDir
	
	if "%x%"=="uf" goto :incomplete
	if "%x:~0,2%"=="uf" (
		echo Opening \urlfiles in \\prdidolspider%x:~2%...
		start \\prdidolspider%x:~2%\urlfiles$
	)

	:: ---Data/Housekeeping------------------------------------------------
	
	if "%x%"=="x" (
		if %HarvestSheetBackup%==yes call :BackupX
		if %HarvestSheetBackup%==ask call :BackupXask
		echo.
		echo Opening %HarvestSheet_File%...
		echo  ^| If opening fails, restart Citrix and try again.
		echo  ^| If the file is being used by another process, ask or wait for the person
		echo  ^| using the file to close it, or open it normally in Windows Explorer or 
		echo  ^| Excel to access it in read-only mode.
		call :excel "%IMEngineers_Path%\%HarvestSheet_File%"
	)
	if "%x%"=="lx" call :LastX
	if "%x%"=="bx" call :BackupX
	if "%x%"=="lj" (
		echo Reading 5 most-recently-modified cfg backup files in %CFGBackup_Path%...
		cd /d %rundir%
		python listjob.py
		if exist out.txt ( 
			start notepad++ out.txt
		) else (
			echo No output file found.
		)
	)
	if "%x%"=="dh" call :DelHTTPC
	if "%x%"=="du" call :DelUFF
	if "%x%"=="da" (
		call :DelHTTPC
		call :DelUFF
	)
	:: ---Other------------------------------------------------------------
	
	if "%x%"=="fs" (
		set x=fs 
		goto :incomplete
	)
	if "%x:~0,2%"=="fs" call :FindString %x:~2%
	
	if "%x%"=="p" (
		echo Downloading files from network...
		if exist "%IMEngineers_Path%\Programs" (
			copy "%IMEngineers_Path%\Programs\IMETasks\IMETasks.bat" %rundir%
			copy "%IMEngineers_Path%\Programs\IMETasks\listjob.py" %rundir%
			echo Done!
			echo.
			echo IMETasks will now quit. Update takes effect the next time you run IMETasks.
			echo.
			pause
			exit
		) else (
			echo Download failed. Check if the network drive is available and if the settings
			echo in your IMETasks.ini file are correct.
		)
	)
	
	if "%x%"=="q" goto :EOF
goto :MAIN

:incomplete
	echo %x%# command incomplete.
goto :MAIN

:UFFPAths
	if %1 geq 9 (
		echo Opening %2 in \\prdidolspider%1...
		start %3 \\prdidolspider%1\urlfilefilter$\%2
		goto :EOF
	)
	if %1 leq 4 (
		echo Opening %2 in \\prdidolspider%1...
		start %3 \\prdidolspider%1\urlfilefilter$\%2
	)
	:: different path in S5-S8
	if %1 geq 5 (
		echo Opening %2 in \\prdidolspider%1\...\urlfilefilter2...
		start %3 \\prdidolspider%1\urlfilefilter$\urlfilefilter2\%2
	)
goto :EOF

:BackupXask
	set y=
	set /p y=Backup first? Enter 'y' if yes: 
	if "%y%"=="y" call :BackupX
goto :EOF

:LastX
	cd /d "%HarvestSheetBackup_Path%"
	if not exist *-*.xlsx (
		echo No backups made.
		goto :MAIN
	)
	for /f "usebackq delims=" %%i in (`dir /b /od *-*.xlsx`) do set file=%%i
	echo Opening %file%...
	call :excel "%file%"
goto :EOF

:BackupX
	cd /d "%IMEngineers_Path%"
	
	set time0=%time:~0,2%%time:~3,2%
	if "%time0:~0,1%"==" " set time0=0%time0:~1%
	set file=x
	if "%date:~0,2%"=="20" set file=%date:~2,2%%date:~5,2%%date:~8,2%-%time0% %HarvestSheet_File%
	if "%date:~10,2%"=="20" set file=%date:~12,2%%date:~7,2%%date:~4,2%-%time0% %HarvestSheet_File%
	if "%file%"=="x" (
		echo Cannot backup. Contact %contact% for assistance.
		goto :EOF
	)
	
	type backups\lastsave.txt
	echo Last backup: %file% by %USERNAME% >> backups\lastsave.txt
	echo New backup:  %file% by %USERNAME%
	
	echo Saving local backup copy...
	copy "%HarvestSheet_File%" "%HarvestSheetBackup_Path%\%file%"
	echo Saving network backup copy...
	copy "%HarvestSheet_File%" "backups\%file%"
	echo Done.
	
	echo Deleting oldest network backup copy...
	cd backups
	call :BackupXdel
	echo Deleting oldest local backup copy...
	cd /d "%HarvestSheetBackup_Path%"
	call :BackupXdel
	echo Done.
	
goto :EOF

:BackupXdel
	:: delete oldest backup file
	set count=0
	for /f "usebackq delims=" %%i in (`dir /b /o-d *-*.xlsx`) do (
		set file=%%i
		set /a count+=1
	)
	if /i %count% geq 10 del "%file%"
goto :EOF

:DelHTTPC
	echo Stopping HTTPConnector...
	net stop HTTPConnector
	
	echo Emptying IDX folder...
	if exist %IDX_Path% (
		cd %IDX_Path%
		del /q *.idx
	)
	
	echo Emptying LOGS folder...
	if exist %Logs_Path% (
		cd %Logs_Path%
		del /q *.log
		rem. > spider.log
	)
	
	echo Deleting SITES folder...
	rd /s/q %Temp_Path%
	
	echo Deleting *.STATUS files...
	del /q %HTTPConnector_Path%\*.status
	
	echo Done.
goto :EOF

:DelUFF
	echo Deleting URLFiles...
	rd /s/q %URLFiles_Path%
	md %URLFiles_Path%
	del /q %URLFileFilter_Path%\ErrorLogDir\*.txt
	del /q %URLFileFilter_Path%\WorkingFolder\*.txt
	echo Done.
goto :EOF

:FindString
	cd /d %CFGBackup_Path%
	echo ---------- filename: occurrences of "%1"
	for /f "usebackq tokens=1-6 delims=\:" %%i in (`find /c /i "%1" *.cfg`) do (
		if not "%%j"==" 0" echo %%i:%%j
	)
goto :EOF

:excel
	:: use local Excel if available
	if "%excel%"=="no" (
		%1
	) else (
		start excel %1
	)
goto :EOF

:: EOF = End Of File
:: "goto :EOF" exits a subroutine