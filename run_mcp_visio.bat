@echo off
echo ===================================================
echo MCP Visio Server - All-in-One Setup
echo ===================================================

setlocal enabledelayedexpansion
set PORT=8051
set DOCKER_PORT=8050

REM Check command line arguments
if "%1"=="relay" goto :start_relay
if "%1"=="docker" goto :start_docker
if "%1"=="test" goto :run_test
if "%1"=="stop" goto :stop_all

REM No arguments, show help menu
if "%1"=="" (
    echo Select an option:
    echo.
    echo 1. Start everything (recommended)
    echo 2. Start Visio relay only
    echo 3. Start MCP Docker container only
    echo 4. Run tests
    echo 5. Stop all services
    echo 6. Exit
    echo.
    
    set /p choice=Enter your choice (1-6): 
    
    if "!choice!"=="1" goto :start_all
    if "!choice!"=="2" goto :start_relay
    if "!choice!"=="3" goto :start_docker
    if "!choice!"=="4" goto :run_test
    if "!choice!"=="5" goto :stop_all
    if "!choice!"=="6" goto :end
    
    echo Invalid choice. Please try again.
    goto :eof
)

:start_all
echo Starting complete MCP Visio environment...
start "Visio Relay" cmd /c "%~f0 relay"
timeout /t 3 /nobreak > nul
call :start_docker
echo.
echo All services started successfully!
echo Visio Relay: http://localhost:%PORT%
echo MCP Server: http://localhost:%DOCKER_PORT%
echo.
echo Use this URL in your applications: http://localhost:%DOCKER_PORT%/mcp
goto :end

:start_relay
echo.
echo Starting Visio Relay Service...
echo This service allows the MCP container to communicate with Microsoft Visio.
echo Make sure Visio is installed on this machine.
echo.
echo Checking for required Python packages...
python -m pip install fastapi uvicorn pywin32 pydantic --quiet
echo Starting Visio Relay Service on port %PORT%...
echo.
echo Press Ctrl+C to stop the service
echo.
python host_visio_relay.py
goto :end

:start_docker
echo.
echo Building and starting MCP Visio Docker container...
docker stop mcp-visio 2>nul
docker rm mcp-visio 2>nul
docker build -t mcp-visio-windows -f Dockerfile.windows .
docker run -d --name mcp-visio ^
  -p %DOCKER_PORT%:%DOCKER_PORT% ^
  -e TRANSPORT=sse ^
  -e HOST=0.0.0.0 ^
  -e PORT=%DOCKER_PORT% ^
  -e VISIO_SERVICE_HOST=host.docker.internal ^
  -e VISIO_SERVICE_PORT=%PORT% ^
  --isolation=process ^
  -v %CD%\src:C:\app\src ^
  -v %CD%\visio-files:C:\visio-files ^
  mcp-visio-windows
echo.
echo Docker container started!
goto :eof

:run_test
echo.
echo Running Visio connection tests...
python test_visio_connection.py
goto :end

:stop_all
echo.
echo Stopping all MCP Visio services...
docker stop mcp-visio 2>nul
docker rm mcp-visio 2>nul
echo.
echo Stopped. You may need to manually close any Visio Relay terminal windows.
goto :end

:end
echo. 