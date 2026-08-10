@echo off
setlocal enabledelayedexpansion

if "%~1"=="" (
    echo Drag and drop one or more MP4 files onto this icon.
    echo.
    pause
    exit /b 0
)

rem ---- Relaunch ourselves hidden (no console) via the sibling .vbs helper ----
if not "%MP4WAV_HIDDEN%"=="1" (
    set "HELPER=%~dp0_mp4towav_hide.vbs"
    if not exist "!HELPER!" (
        echo Hidden-launch helper not found: !HELPER!
        echo Running with the console visible instead.
        echo.
        goto :run_visible
    )
    set MP4WAV_HIDDEN=1
    start "" wscript.exe "!HELPER!" "%~f0" %*
    exit /b 0
)

rem ============================================================
rem  Hidden execution path - no console output, logs to %TEMP%
rem ============================================================
set "FFMPEG=ffmpeg"
where %FFMPEG% >nul 2>&1
if errorlevel 1 set "FFMPEG=C:\ffmpeg-master-latest-win64-gpl-shared\bin\ffmpeg.exe"

set "FAILED=0"
set "LOG=%TEMP%\mp4towav.log"
> "%LOG%" echo MP4 to WAV - %DATE% %TIME%

:hloop
if "%~1"=="" goto :hdone
set "IN=%~1"
set "OUT=%~dpn1.wav"
>>"%LOG%" echo.
>>"%LOG%" echo === %~nx1
"%FFMPEG%" -hide_banner -y -i "%IN%" -vn -acodec pcm_s16le "%OUT%" >>"%LOG%" 2>&1
if errorlevel 1 (
    set /a FAILED+=1
    >>"%LOG%" echo *** FAILED: %~nx1
)
shift
goto :hloop

:hdone
if !FAILED! gtr 0 (
    mshta "vbscript:msgbox(""MP4 to WAV: !FAILED! file(s) failed. See log at %LOG%."",48,""MP4 to WAV""):close"
)
endlocal
exit /b 0

rem ============================================================
rem  Visible execution path - fallback if the .vbs helper is gone
rem ============================================================
:run_visible
set "FFMPEG=ffmpeg"
where %FFMPEG% >nul 2>&1
if errorlevel 1 set "FFMPEG=C:\ffmpeg-master-latest-win64-gpl-shared\bin\ffmpeg.exe"
set "FAILED=0"

:vloop
if "%~1"=="" goto :vdone
set "IN=%~1"
set "OUT=%~dpn1.wav"
echo.
echo === Extracting: %~nx1
"%FFMPEG%" -hide_banner -y -i "%IN%" -vn -acodec pcm_s16le "%OUT%"
if errorlevel 1 (
    echo *** FAILED: %~nx1
    set /a FAILED+=1
) else (
    echo Saved: %OUT%
)
shift
goto :vloop

:vdone
echo.
if !FAILED! gtr 0 (
    echo Done with !FAILED! failure^(s^).
    pause
) else (
    echo Done.
    timeout /t 2 >nul
)
endlocal
exit /b 0
