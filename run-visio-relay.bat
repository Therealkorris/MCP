@echo off
echo Starting Visio Relay Service...

REM Check if Python is installed
python --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
  echo ERROR: Python is not installed. Please install Python 3.x.
  exit /b 1
)

REM Install required packages if needed
echo Installing required packages...
pip install fastapi uvicorn pywin32 pydantic >nul 2>&1

REM Run the relay service
echo Starting Visio Relay Service on port 8051...
start "Visio Relay Service" python visio_relay.py

echo.
echo Visio Relay Service is starting in a new window.
echo The service will be available at http://localhost:8051
echo This service allows Docker containers to interact with Visio on your host machine.
echo.
echo Use Ctrl+C in the service window to stop the service when you're done.
echo. 