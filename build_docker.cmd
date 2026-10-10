@echo off
cd /d .
echo Building qa-auto-mix Docker image...
echo Start time: %date% %time%
docker build -t qa-auto-mix:latest -f Dockerfile . 2>&1
echo.
echo Build exit code: %ERRORLEVEL%
echo End time: %date% %time%
if %ERRORLEVEL% EQU 0 (
    echo SUCCESS - Image built!
) else (
    echo FAILED
)
pause
